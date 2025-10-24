"""
Email Mass Sender - Main Application
Professional web interface for managing email campaigns and accounts
"""

import os
import asyncio
import time
import sys

# Windows-specific asyncio policy to prevent RuntimeError
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, session, Response
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, login_user, logout_user, login_required, current_user, UserMixin
from werkzeug.utils import secure_filename
import json
from datetime import datetime, timedelta
from loguru import logger
from email_validator import validate_email, EmailNotValidError
from typing import Any
from utils.dynamic_timing import DynamicTiming, TimingContext

# Custom JSON serializer for datetime objects
def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable")

def safe_async_run(coro):
    """Safely run async coroutine with proper Windows asyncio handling"""
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            # If loop is already running, use ThreadPoolExecutor
            import concurrent.futures
            with concurrent.futures.ThreadPoolExecutor() as executor:
                future = executor.submit(asyncio.run, coro)
                return future.result()
        else:
            return asyncio.run(coro)
    except RuntimeError:
        # Fallback: run in new event loop
        return asyncio.run(coro)

# Import our modules
from config import config, init_database
from database import EmailAccount, EmailCampaign, EmailRecipient, EmailLog, get_db
from utils.progress_tracker import init_progress_tracker, get_progress_tracker
from email_service.sender import EmailSender, EmailTask
from ai.placeholder_replacer import placeholder_replacer
from auth.base_auth import TokenManager
from auth.office365_auth import Office365Auth
from auth.gmail_auth import GmailAuth
from auth.yahoo_auth import YahooAuth
from auth.hotmail_auth import HotmailAuth
from automation.office365 import login as o365_login, send_with_bcc as o365_send
from automation.office365_fast import fast_send_email, fast_send_email_with_cookies
from email_service.cookie_manager import cookie_manager

# Note: We do not override the Windows event loop policy globally because
# Playwright requires Proactor event loop support for subprocess handling.

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = config.SECRET_KEY
app.config['SQLALCHEMY_DATABASE_URI'] = config.DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = config.UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = config.MAX_ATTACHMENT_SIZE

# Initialize extensions
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Initialize database
db_manager = init_database(config.DATABASE_URL)
email_sender = EmailSender(db_manager)

# Initialize progress tracker
progress_tracker = init_progress_tracker(db_manager)

# Create upload directory
os.makedirs(config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs('logs', exist_ok=True)
os.makedirs('sessions', exist_ok=True)

# Configure logging
logger.add(
    config.LOG_FILE,
    rotation="1 day",
    retention="30 days",
    level=config.LOG_LEVEL,
    format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}"
)

# Add console logging for real-time output
logger.add(
    lambda msg: print(msg, end=""),
    level=config.LOG_LEVEL,
    format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}"
)

class User(UserMixin):
    """Minimal user object for Flask-Login"""
    def __init__(self, user_id: int, username: str):
        self.id = str(user_id)
        self.username = username

    @property
    def is_active(self):
        return True

    @property
    def is_authenticated(self):
        return True

    @property
    def is_anonymous(self):
        return False


@login_manager.user_loader
def load_user(user_id):
    """Load user for Flask-Login"""
    # Single admin user for demo; replace with real user lookup in production
    return User(user_id, 'admin')

@app.route('/')
@login_required
def dashboard():
    """Simplified dashboard focused on direct send"""
    try:
        # Get basic stats without campaign complexity
        active_accounts = db_manager.get_active_accounts()
        active_accounts_count = len(active_accounts)
        
        # Simple stats (no campaign tracking needed)
        total_sent_count = 0  # Could be tracked separately if needed
        success_rate = 100    # Default success rate
        failed_count = 0      # Could be tracked separately if needed
        
        return render_template('dashboard_simple.html',
                             active_accounts_count=active_accounts_count,
                             total_sent_count=total_sent_count,
                             success_rate=success_rate,
                             failed_count=failed_count)
    except Exception as e:
        logger.error(f"Dashboard error: {e}")
        return render_template('dashboard_simple.html',
                             active_accounts_count=0,
                             total_sent_count=0,
                             success_rate=0,
                             failed_count=0)

@app.route('/accounts')
@login_required
def accounts():
    """Account management page"""
    try:
        # Get regular database accounts
        db_accounts = db_manager.get_active_accounts()
        
        # Get cookie-based Office 365 accounts
        cookie_accounts = cookie_manager.get_all_accounts()
        
        # Convert cookie accounts to a format compatible with the template
        cookie_accounts_list = []
        for email, account_data in cookie_accounts.items():
            # Create a mock account object for template compatibility
            class MockAccount:
                def __init__(self, email, provider, account_type, status, added_at, last_used):
                    self.email = email
                    self.provider = provider
                    self.account_type = account_type
                    self.status = status
                    self.added_at = added_at
                    self.last_used = last_used
                    self.is_active = status == 'active'
                    self.is_verified = True
                    self.success_count = 0
                    self.error_count = 0
                    self.id = f"cookie_{email}"
                    # Add missing attributes for template compatibility
                    self.created_at = datetime.strptime(added_at, '%Y-%m-%dT%H:%M:%S.%f') if added_at else datetime.utcnow()
                    self.updated_at = datetime.utcnow()
            
            cookie_accounts_list.append(MockAccount(
                email=email,
                provider='office365',
                account_type=account_data.get('account_type', 'paid'),
                status=account_data.get('status', 'active'),
                added_at=account_data.get('added_at', ''),
                last_used=account_data.get('last_used', '')
            ))
        
        # Combine both account types
        all_accounts = list(db_accounts) + cookie_accounts_list
        
        return render_template('accounts.html', accounts=all_accounts, cookie_accounts=cookie_accounts)
    except Exception as e:
        logger.error(f"Accounts page error: {e}")
        flash('Error loading accounts', 'error')
        return render_template('accounts.html', accounts=[], cookie_accounts={})

@app.route('/accounts/active')
@login_required
def get_active_accounts():
    """Return combined list of active sending accounts for Direct Send UI"""
    try:
        db_accounts = db_manager.get_active_accounts() or []
        cookie_accounts = cookie_manager.get_all_accounts() or {}

        combined = []
        for acc in db_accounts:
            email = getattr(acc, 'email', None)
            if not email:
                continue
            combined.append({
                'email': email,
                'provider': getattr(acc, 'provider', 'office365'),
                'status': 'active',
                'type': 'database'
            })

        for email, data in cookie_accounts.items():
            if not email:
                continue
            combined.append({
                'email': email,
                'provider': 'office365',
                'status': data.get('status', 'active'),
                'type': 'cookie'
            })

        # Deduplicate by email
        dedup = {}
        for item in combined:
            dedup[item['email']] = item

        return jsonify({'success': True, 'accounts': list(dedup.values())})
    except Exception as e:
        logger.error(f"Get active accounts error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/accounts/<int:account_id>/details')
@login_required
def account_details(account_id):
    """Return HTML fragment with account details"""
    try:
        session_db = db_manager.get_session()
        account = session_db.query(EmailAccount).filter(EmailAccount.id == account_id).first()
        if not account:
            return jsonify({'success': False, 'error': 'Account not found'})

        # Build a simple HTML snippet (kept minimal; rendered in modal)
        html = f"""
        <div class=\"row\">
            <div class=\"col-md-6\">
                <ul class=\"list-group list-group-flush\">
                    <li class=\"list-group-item\"><strong>Email:</strong> {account.email}</li>
                    <li class=\"list-group-item\"><strong>Provider:</strong> {account.provider}</li>
                    <li class=\"list-group-item\"><strong>Status:</strong> {'Active' if account.is_active else 'Inactive'}</li>
                    <li class=\"list-group-item\"><strong>Verified:</strong> {'Yes' if account.is_verified else 'No'}</li>
                </ul>
            </div>
            <div class=\"col-md-6\">
                <ul class=\"list-group list-group-flush\">
                    <li class=\"list-group-item\"><strong>Success count:</strong> {account.success_count}</li>
                    <li class=\"list-group-item\"><strong>Error count:</strong> {account.error_count}</li>
                    <li class=\"list-group-item\"><strong>Last used:</strong> {account.last_used.strftime('%Y-%m-%d %H:%M') if account.last_used else 'Never'}</li>
                    <li class=\"list-group-item\"><strong>Created:</strong> {account.created_at.strftime('%Y-%m-%d')}</li>
                </ul>
            </div>
        </div>
        """
        return jsonify({'success': True, 'html': html})
    except Exception as e:
        logger.error(f"Account details error: {e}")
        return jsonify({'success': False, 'error': str(e)})

def _get_provider_for_name(provider: str):
    providers = {
        'office365': Office365Auth(),
        'gmail': GmailAuth(),
        'yahoo': YahooAuth(),
        'hotmail': HotmailAuth()
    }
    return providers.get(provider)

@app.route('/accounts/<int:account_id>/test', methods=['POST'])
@login_required
def test_account(account_id):
    """Test the connection for a single account"""
    try:
        session_db = db_manager.get_session()
        account = session_db.query(EmailAccount).filter(EmailAccount.id == account_id).first()
        if not account:
            return jsonify({'success': False, 'error': 'Account not found'})

        provider = _get_provider_for_name(account.provider)
        if not provider:
            return jsonify({'success': False, 'error': f'Unsupported provider: {account.provider}'})

        # Try to obtain a valid token if the account uses OAuth
        is_ok = False
        try:
            token = asyncio.run(TokenManager(db_manager).get_valid_token(account.id, account.provider))
            user_info = asyncio.run(provider.get_user_info(token))
            is_ok = bool(user_info.get('email'))
        except Exception as e:
            # Fallback for Office365 browser-automation accounts without tokens: check session file
            if account.provider == 'office365':
                try:
                    from automation.office365_fast import _storage_path  # type: ignore
                    storage = _storage_path(account.email)
                    is_ok = storage.exists()
                except Exception:
                    is_ok = False
            else:
                logger.warning(f"Account test token check failed for {account.email}: {e}")

        if is_ok:
            account.is_verified = True
            account.last_used = datetime.utcnow()
            session_db.commit()
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': 'Connection test failed'})
    except Exception as e:
        logger.error(f"Test account error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/accounts/test-all', methods=['POST'])
@login_required
def test_all_accounts():
    """Test all active accounts"""
    try:
        session_db = db_manager.get_session()
        accounts = session_db.query(EmailAccount).filter(EmailAccount.is_active == True).all()
        tested = 0
        successful = 0
        failed = 0
        for acc in accounts:
            tested += 1
            try:
                resp = test_account(acc.id)
                # When calling function directly inside same request, resp is a Response
                # Extract JSON payload safely
                if hasattr(resp, 'get_json'):
                    data = resp.get_json() or {}
                    ok = bool(data.get('success'))
                else:
                    ok = False
            except Exception:
                ok = False
            if ok:
                successful += 1
            else:
                failed += 1
        return jsonify({'success': True, 'tested': tested, 'successful': successful, 'failed': failed})
    except Exception as e:
        logger.error(f"Test all accounts error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/accounts/<int:account_id>/refresh', methods=['POST'])
@login_required
def refresh_account(account_id):
    """Refresh OAuth token for an account"""
    try:
        session_db = db_manager.get_session()
        account = session_db.query(EmailAccount).filter(EmailAccount.id == account_id).first()
        if not account:
            return jsonify({'success': False, 'error': 'Account not found'})
        if not account.refresh_token:
            return jsonify({'success': False, 'error': 'No refresh token stored for this account'})
        provider = _get_provider_for_name(account.provider)
        if not provider:
            return jsonify({'success': False, 'error': f'Unsupported provider: {account.provider}'})

        tokens = asyncio.run(provider.refresh_access_token(account.refresh_token))
        db_manager.update_account_tokens(
            account_id=account.id,
            access_token=tokens['access_token'],
            refresh_token=tokens.get('refresh_token'),
            expires_at=tokens['expires_at']
        )
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f"Refresh account error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/accounts/<int:account_id>/delete', methods=['DELETE'])
@login_required
def delete_account(account_id):
    """Delete an account and related data where safe"""
    try:
        session_db = db_manager.get_session()
        account = session_db.query(EmailAccount).filter(EmailAccount.id == account_id).first()
        if not account:
            return jsonify({'success': False, 'error': 'Account not found'})
        session_db.delete(account)
        session_db.commit()
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f"Delete account error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/accounts/add')
@login_required
def add_account():
    """Add new account page"""
    return render_template('add_account.html')

@app.route('/accounts/oauth/<provider>')
@login_required
def oauth_redirect(provider):
    """OAuth redirect for account authentication"""
    try:
        # Get provider authentication instance
        providers = {
            'office365': Office365Auth(),
            'gmail': GmailAuth(),
            'yahoo': YahooAuth(),
            'hotmail': HotmailAuth()
        }
        
        if provider not in providers:
            flash(f'Unsupported provider: {provider}', 'error')
            return redirect(url_for('add_account'))
        
        auth_provider = providers[provider]
        auth_url = auth_provider.get_auth_url()
        
        # Store provider in session for callback
        session['oauth_provider'] = provider
        
        return redirect(auth_url)
    
    except Exception as e:
        logger.error(f"OAuth redirect error: {e}")
        flash('Error initiating OAuth authentication', 'error')
        return redirect(url_for('add_account'))

@app.route('/callback/<provider>')
@login_required
def oauth_callback(provider):
    """OAuth callback handler"""
    try:
        code = request.args.get('code')
        state = request.args.get('state')
        error = request.args.get('error')
        
        if error:
            flash(f'OAuth error: {error}', 'error')
            return redirect(url_for('add_account'))
        
        if not code:
            flash('No authorization code received', 'error')
            return redirect(url_for('add_account'))
        
        # Get provider authentication instance
        providers = {
            'office365': Office365Auth(),
            'gmail': GmailAuth(),
            'yahoo': YahooAuth(),
            'hotmail': HotmailAuth()
        }
        
        if provider not in providers:
            flash(f'Unsupported provider: {provider}', 'error')
            return redirect(url_for('add_account'))
        
        auth_provider = providers[provider]
        
        # Exchange code for tokens (run in async context)
        import asyncio
        tokens = asyncio.run(auth_provider.exchange_code_for_token(code, state))
        
        # Get user info (run in async context)
        user_info = asyncio.run(auth_provider.get_user_info(tokens['access_token']))
        
        # Add account to database
        account = db_manager.add_email_account(
            email=user_info['email'],
            provider=provider,
            access_token=tokens['access_token'],
            refresh_token=tokens.get('refresh_token'),
            token_expires_at=tokens['expires_at']
        )
        
        flash(f'Account {user_info["email"]} added successfully!', 'success')
        return redirect(url_for('accounts'))
    
    except Exception as e:
        logger.error(f"OAuth callback error: {e}")
        flash(f'Error adding account: {str(e)}', 'error')
        return redirect(url_for('add_account'))

@app.route('/campaigns')
@login_required
def campaigns():
    """Campaign management page"""
    try:
        campaigns = db_manager.get_session().query(EmailCampaign).order_by(
            EmailCampaign.created_at.desc()
        ).all()
        return render_template('campaigns.html', campaigns=campaigns)
    except Exception as e:
        logger.error(f"Campaigns page error: {e}")
        flash('Error loading campaigns', 'error')
        return render_template('campaigns.html', campaigns=[])

@app.route('/campaigns/new')
@login_required
def new_campaign():
    """New campaign page"""
    try:
        defaults = db_manager.get_campaign_defaults()
    except Exception:
        defaults = {'subject': '', 'body': ''}
    return render_template('new_campaign.html', default_subject=defaults.get('subject', ''), default_body=defaults.get('body', ''))

@app.route('/campaigns/create', methods=['POST'])
@login_required
def create_campaign():
    """Create new campaign"""
    try:
        data = request.get_json()
        
        # Validate required fields
        required_fields = ['name', 'subject', 'body_html']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'})
        
        # Create campaign
        campaign = db_manager.create_campaign(
            name=data['name'],
            subject=data['subject'],
            body_html=data['body_html'],
            body_text=data.get('body_text'),
            template_data=data.get('template_data', {})
        )
        # Persist latest subject/body as defaults for New Campaign form
        try:
            db_manager.set_campaign_defaults(data['subject'], data['body_html'])
        except Exception as persist_err:
            logger.debug(f"Failed to persist campaign defaults: {persist_err}")
        
        return jsonify({
            'success': True,
            'campaign_id': campaign.id,
            'message': 'Campaign created successfully'
        })
    
    except Exception as e:
        logger.error(f"Create campaign error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>/recipients', methods=['POST'])
@login_required
def add_recipients(campaign_id):
    """Add recipients to campaign"""
    try:
        data = request.get_json()
        recipients = data.get('recipients', [])
        
        if not recipients:
            return jsonify({'success': False, 'error': 'No recipients provided'})
        
        # Add recipients to database
        db_manager.add_recipients(campaign_id, recipients)
        
        return jsonify({
            'success': True,
            'message': f'Added {len(recipients)} recipients to campaign'
        })
    
    except Exception as e:
        logger.error(f"Add recipients error: {e}")
        return jsonify({'success': False, 'error': str(e)})

# ============================================================================
# PROGRESS TRACKING ROUTES
# ============================================================================

@app.route('/progress/<session_id>')
@login_required
def get_progress(session_id):
    """Get current progress for a session"""
    try:
        tracker = get_progress_tracker()
        progress = tracker.get_progress(session_id)
        
        if progress:
            return jsonify({'success': True, 'progress': progress})
        else:
            return jsonify({'success': False, 'error': 'Session not found'})
            
    except Exception as e:
        logger.error(f"Error getting progress: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/progress/<session_id>/stream')
@login_required
def progress_stream(session_id):
    """Server-Sent Events stream for real-time progress updates"""
    def generate():
        tracker = get_progress_tracker()
        last_progress = None
        last_status = None
        update_count = 0
        
        logger.info(f"Starting progress stream for session {session_id}")
        
        while True:
            try:
                progress = tracker.get_progress(session_id)
                
                if not progress:
                    logger.warning(f"Session {session_id} not found in progress stream")
                    yield f"data: {json.dumps({'error': 'Session not found'})}\n\n"
                    break
                
                current_status = progress.get('status', 'unknown')
                current_progress = progress.get('progress_percentage', 0)
                sent_count = progress.get('sent_count', 0)
                failed_count = progress.get('failed_count', 0)
                
                # Always send update if status changed
                status_changed = current_status != last_status
                
                # Send update if progress changed significantly or status changed
                progress_changed = (
                    progress != last_progress or 
                    status_changed or
                    update_count < 3  # Send first few updates regardless
                )
                
                if progress_changed:
                    logger.info(f"Progress update {update_count + 1}: Status={current_status}, Progress={current_progress:.1f}%, Sent={sent_count}, Failed={failed_count}")
                    yield f"data: {json.dumps({'progress': progress}, default=json_serial)}\n\n"
                    last_progress = progress.copy()
                    last_status = current_status
                    update_count += 1
                
                # Check if completed
                if current_status in ['completed', 'failed', 'cancelled']:
                    logger.info(f"Session {session_id} completed with status: {current_status}")
                    yield f"data: {json.dumps({'progress': progress, 'completed': True}, default=json_serial)}\n\n"
                    break
                
                # Dynamic sleep based on progress rate
                sleep_time = 0.5 if current_progress > 0.8 else 1.0  # Faster updates near completion
                time.sleep(sleep_time)
                
            except Exception as e:
                logger.error(f"Error in progress stream for session {session_id}: {e}")
                yield f"data: {json.dumps({'error': str(e)}, default=json_serial)}\n\n"
                break
        
        logger.info(f"Progress stream ended for session {session_id} after {update_count} updates")
    
    return Response(generate(), mimetype='text/event-stream', headers={
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Cache-Control'
    })

# =========================================================================
# OFFICE365 EXPORT: Expanded People Data per account
# =========================================================================

@app.route('/office365/accounts/<path:email>/export-expanded-people', methods=['GET'])
@login_required
def export_office365_expanded_people(email):
    """Download Outlook 'expanded people data' using existing cookie session.

    Returns a text/plain file named: export_from_<email>.txt
    """
    try:
        logger.info(f"[EXPORT] Start export for {email}")
        from pathlib import Path as _Path
        account = cookie_manager.get_account(email)
        logger.info(f"[EXPORT] Cookie account present: {bool(account)}")

        from playwright.sync_api import sync_playwright
        from automation.office365_fast import _outlook_host_for_email, _inject_cookies_to_context, _storage_path

        filename = f"export_from_{email}.txt"

        with sync_playwright() as p:
            logger.info("[EXPORT] Launching Chromium (headless)")
            browser = p.chromium.launch(headless=True)

            # Prefer stored Playwright session (same as sending flow); otherwise try cookies
            context = None
            outlook_host = _outlook_host_for_email(email)

            # Primary: saved storage_state
            storage_file = _storage_path(email)
            logger.info(f"[EXPORT] Storage file: {storage_file} (exists={storage_file.exists()})")
            if storage_file.exists():
                logger.info("[EXPORT] Using storage_state session")
                context = browser.new_context(accept_downloads=True, storage_state=str(storage_file))
            elif account and (account.get('cookies') or []):
                logger.info("[EXPORT] Using cookie injection fallback")
                # Fallback: inject cookies
                context = browser.new_context(accept_downloads=True)
                cookies = account.get('cookies') or []
                _inject_cookies_to_context(context, cookies, f'.{outlook_host}')
            else:
                logger.error("[EXPORT] No session or cookies available")
                browser.close()
                return jsonify({'success': False, 'error': f'No saved session or cookies found for {email}. Please log in once from automation login.'}), 404

            page = context.new_page()
            # First go to mailbox to ensure auth context, then navigate to export settings
            mailbox_url = f"https://{outlook_host}/mail/"
            logger.info(f"[EXPORT] Navigate: {mailbox_url}")
            page.goto(mailbox_url, wait_until="domcontentloaded")
            # Dynamic wait for page load
            if not DynamicTiming.wait_for_page_load(page, "mail", max_wait=3000):
                logger.warning("Mailbox page load timeout")
            # Try direct settings URL first; if auth kicks in, fall back to in-UI navigation
            # Outlook host differs for consumer vs work; path is similar
            settings_url = f"https://{outlook_host}/mail/options/general/export"
            logger.info(f"[EXPORT] Navigate: {settings_url}")
            page.goto(settings_url, wait_until="domcontentloaded")
            try:
                page.wait_for_load_state('networkidle', timeout=15000)
            except Exception:
                pass
            # Dynamic wait for settings page load
            if not DynamicTiming.wait_for_page_load(page, "export", max_wait=3000):
                logger.warning("Settings page load timeout")

            if 'login' in page.url.lower() or 'signin' in page.url.lower():
                logger.warning(f"[EXPORT] Direct URL triggered login; trying in-UI settings navigation")
                # Go back to mailbox and open settings panel via UI
                page.goto(mailbox_url, wait_until="domcontentloaded")
                try:
                    page.wait_for_load_state('networkidle', timeout=15000)
                except Exception:
                    pass
                # Dynamic wait for settings panel to open
                if not DynamicTiming.wait_for_condition(
                    lambda: page.locator('[data-testid*="settings"], [aria-label*="Settings"]').count() > 0,
                    max_wait=2000
                ):
                    logger.warning("Settings panel not found")
                
                # Click the Settings (gear) button
                gear_selectors = [
                    "button[aria-label='Settings']",
                    "button[title='Settings']",
                    "button:has([data-icon-name='Settings'])",
                    "button:has-text('Settings')"
                ]
                clicked_gear = False
                for gs in gear_selectors:
                    try:
                        logger.info(f"[EXPORT] Try open Settings with: {gs}")
                        btn = page.locator(gs).first
                        btn.wait_for(state='visible', timeout=5000)
                        btn.click()
                        clicked_gear = True
                        break
                    except Exception as egear:
                        logger.debug(f"[EXPORT] Settings open failed: {gs} | {egear}")
                        continue
                if not clicked_gear:
                    browser.close()
                    return jsonify({'success': False, 'error': 'Unable to open Settings gear in Outlook UI'}), 500

                # Click "View all Outlook settings"
                view_all_selectors = [
                    "text=View all Outlook settings",
                    "a:has-text('View all Outlook settings')",
                    "button:has-text('View all Outlook settings')"
                ]
                opened_all = False
                for vs in view_all_selectors:
                    try:
                        logger.info(f"[EXPORT] Open 'View all Outlook settings' with: {vs}")
                        item = page.locator(vs).first
                        item.scroll_into_view_if_needed(timeout=3000)
                        item.wait_for(state='visible', timeout=5000)
                        item.click()
                        opened_all = True
                        break
                    except Exception as eopen:
                        logger.debug(f"[EXPORT] View all click failed: {vs} | {eopen}")
                        continue
                if not opened_all:
                    browser.close()
                    return jsonify({'success': False, 'error': "Couldn't open 'View all Outlook settings'"}), 500

                # Navigate General -> Privacy and data
                try:
                    logger.info("[EXPORT] Navigating Settings: General -> Privacy and data")
                    page.locator("role=tab[name='General']").first.click(timeout=5000)
                except Exception:
                    try:
                        page.locator("text=General").first.click(timeout=5000)
                    except Exception:
                        logger.debug("[EXPORT] Could not click General tab; continuing")
                try:
                    page.locator("text=Privacy and data").first.click(timeout=6000)
                except Exception as epriv:
                    logger.debug(f"[EXPORT] Could not click 'Privacy and data': {epriv}")
                
                # Now we are in settings pane UI; proceed to locate the download control later below

            candidates = [
                "text=Download my expanded people data",
                "button:has-text('Download my expanded people data')",
                "a:has-text('Download my expanded people data')",
                "text=Download expanded people data",
                "text=Download my people data"
            ]

            download = None
            logger.info("[EXPORT] Locating download control")
            # 1) Try on main page
            for sel in candidates:
                try:
                    logger.info(f"[EXPORT] Try selector (main): {sel}")
                    loc = page.locator(sel).first
                    loc.scroll_into_view_if_needed(timeout=5000)
                    loc.wait_for(state='visible', timeout=5000)
                    if loc.is_visible():
                        with page.expect_download(timeout=30000) as dl_info:
                            loc.click()
                        download = dl_info.value
                        logger.info("[EXPORT] Download event captured (main)")
                        break
                except Exception as sel_err:
                    logger.debug(f"[EXPORT] Selector failed (main): {sel} | {sel_err}")
                    continue

            # 2) Try within iframes
            if not download:
                try:
                    frames = page.frames
                    logger.info(f"[EXPORT] Frames found: {len(frames)}")
                    for f in frames:
                        try:
                            logger.info(f"[EXPORT] Checking frame url={getattr(f, 'url', '')}")
                        except Exception:
                            pass
                        for sel in candidates:
                            try:
                                logger.info(f"[EXPORT] Try selector (frame): {sel}")
                                floc = f.locator(sel).first
                                floc.scroll_into_view_if_needed(timeout=5000)
                                floc.wait_for(state='visible', timeout=5000)
                                if floc.is_visible():
                                    with page.expect_download(timeout=30000) as dl_info:
                                        floc.click()
                                    download = dl_info.value
                                    logger.info("[EXPORT] Download event captured (frame)")
                                    raise Exception("__EXPORT_FOUND__")
                            except Exception as sel_err2:
                                logger.debug(f"[EXPORT] Selector failed (frame): {sel} | {sel_err2}")
                                continue
                except Exception as breaker:
                    if str(breaker) != "__EXPORT_FOUND__":
                        logger.debug(f"[EXPORT] Frame scan ended: {breaker}")

            if not download:
                try:
                    logger.info("[EXPORT] Trying generic [download] element fallback")
                    with page.expect_download(timeout=10000) as dl_info:
                        page.locator('a[download], button[download]').first.click()
                    download = dl_info.value
                    logger.info("[EXPORT] Download event captured via fallback")
                except Exception as fb_err:
                    logger.error(f"[EXPORT] Fallback failed: {fb_err}")
                    browser.close()
                    return jsonify({'success': False, 'error': 'Could not locate export control on Outlook page.'}), 500

            temp_path = download.path()
            content_bytes = b""
            if temp_path:
                with open(temp_path, 'rb') as f:
                    content_bytes = f.read()
            else:
                content_bytes = download.content()

            browser.close()

            from flask import make_response
            response = make_response(content_bytes)
            response.headers.set('Content-Type', 'text/plain; charset=utf-8')
            response.headers.set('Content-Disposition', f'attachment; filename="{filename}"')
            logger.info(f"[EXPORT] Success. Bytes={len(content_bytes)} File={filename}")
            return response

    except Exception as e:
        logger.exception(f"[EXPORT] Error for {email}: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/office365/accounts/<path:email>/export-processed-contacts', methods=['GET'])
@login_required
def export_processed_contacts(email):
    """Download and process Outlook contacts, returning 3 categorized files.
    
    Returns a ZIP file containing:
    - {email}-office365.txt (Microsoft 365 emails)
    - {email}-gsuite.txt (Google Workspace emails)  
    - {email}-others.txt (Other email providers)
    """
    try:
        logger.info(f"[PROCESSED_EXPORT] Start processed export for {email}")
        
        # Import required modules
        from pathlib import Path as _Path
        import zipfile
        import io
        from utils.email_processor import email_processor
        
        # Get account info
        account = cookie_manager.get_account(email)
        logger.info(f"[PROCESSED_EXPORT] Cookie account present: {bool(account)}")

        from playwright.sync_api import sync_playwright
        from automation.office365_fast import _outlook_host_for_email, _inject_cookies_to_context, _storage_path

        with sync_playwright() as p:
            logger.info("[PROCESSED_EXPORT] Launching Chromium (headless)")
            browser = p.chromium.launch(headless=True)

            # Authentication setup (same as original export)
            context = None
            outlook_host = _outlook_host_for_email(email)

            # Primary: saved storage_state
            storage_file = _storage_path(email)
            logger.info(f"[PROCESSED_EXPORT] Storage file: {storage_file} (exists={storage_file.exists()})")
            if storage_file.exists():
                logger.info("[PROCESSED_EXPORT] Using storage_state session")
                context = browser.new_context(accept_downloads=True, storage_state=str(storage_file))
            elif account and (account.get('cookies') or []):
                logger.info("[PROCESSED_EXPORT] Using cookie injection fallback")
                context = browser.new_context(accept_downloads=True)
                cookies = account.get('cookies') or []
                _inject_cookies_to_context(context, cookies, f'.{outlook_host}')
            else:
                logger.error("[PROCESSED_EXPORT] No session or cookies available")
                browser.close()
                return jsonify({'success': False, 'error': f'No saved session or cookies found for {email}. Please log in once from automation login.'}), 404

            page = context.new_page()
            
            # Navigate to export settings (same logic as original)
            mailbox_url = f"https://{outlook_host}/mail/"
            logger.info(f"[PROCESSED_EXPORT] Navigate: {mailbox_url}")
            page.goto(mailbox_url, wait_until="domcontentloaded")
            
            if not DynamicTiming.wait_for_page_load(page, "mail", max_wait=3000):
                logger.warning("Mailbox page load timeout")
            
            settings_url = f"https://{outlook_host}/mail/options/general/export"
            logger.info(f"[PROCESSED_EXPORT] Navigate: {settings_url}")
            page.goto(settings_url, wait_until="domcontentloaded")
            
            try:
                page.wait_for_load_state('networkidle', timeout=15000)
            except Exception:
                pass
            
            if not DynamicTiming.wait_for_page_load(page, "export", max_wait=3000):
                logger.warning("Settings page load timeout")

            # Handle authentication redirect (same as original)
            if 'login' in page.url.lower() or 'signin' in page.url.lower():
                logger.warning(f"[PROCESSED_EXPORT] Direct URL triggered login; trying in-UI settings navigation")
                page.goto(mailbox_url, wait_until="domcontentloaded")
                
                # UI navigation logic (simplified version of original)
                try:
                    page.wait_for_load_state('networkidle', timeout=15000)
                except Exception:
                    pass
                
                # Click Settings gear
                gear_selectors = [
                    "button[aria-label='Settings']",
                    "button[title='Settings']",
                    "button:has([data-icon-name='Settings'])",
                    "button:has-text('Settings')"
                ]
                
                clicked_gear = False
                for gs in gear_selectors:
                    try:
                        btn = page.locator(gs).first
                        btn.wait_for(state='visible', timeout=5000)
                        btn.click()
                        clicked_gear = True
                        break
                    except Exception:
                        continue
                
                if not clicked_gear:
                    browser.close()
                    return jsonify({'success': False, 'error': 'Unable to open Settings gear in Outlook UI'}), 500

                # Click "View all Outlook settings"
                view_all_selectors = [
                    "text=View all Outlook settings",
                    "a:has-text('View all Outlook settings')",
                    "button:has-text('View all Outlook settings')"
                ]
                
                opened_all = False
                for vs in view_all_selectors:
                    try:
                        item = page.locator(vs).first
                        item.scroll_into_view_if_needed(timeout=3000)
                        item.wait_for(state='visible', timeout=5000)
                        item.click()
                        opened_all = True
                        break
                    except Exception:
                        continue
                
                if not opened_all:
                    browser.close()
                    return jsonify({'success': False, 'error': "Couldn't open 'View all Outlook settings'"}), 500

                # Navigate to General -> Privacy and data
                try:
                    page.locator("role=tab[name='General']").first.click(timeout=5000)
                except Exception:
                    try:
                        page.locator("text=General").first.click(timeout=5000)
                    except Exception:
                        pass
                
                try:
                    page.locator("text=Privacy and data").first.click(timeout=6000)
                except Exception:
                    pass

            # Find and click export control
            candidates = [
                "text=Download my expanded people data",
                "button:has-text('Download my expanded people data')",
                "a:has-text('Download my expanded people data')",
                "text=Download expanded people data",
                "text=Download my people data"
            ]

            download = None
            logger.info("[PROCESSED_EXPORT] Locating download control")
            
            # Try main page first
            for sel in candidates:
                try:
                    loc = page.locator(sel).first
                    loc.scroll_into_view_if_needed(timeout=5000)
                    loc.wait_for(state='visible', timeout=5000)
                    if loc.is_visible():
                        with page.expect_download(timeout=30000) as dl_info:
                            loc.click()
                        download = dl_info.value
                        logger.info("[PROCESSED_EXPORT] Download event captured (main)")
                        break
                except Exception:
                    continue

            # Try iframes if not found on main page
            if not download:
                try:
                    frames = page.frames
                    for f in frames:
                        for sel in candidates:
                            try:
                                floc = f.locator(sel).first
                                floc.scroll_into_view_if_needed(timeout=5000)
                                floc.wait_for(state='visible', timeout=5000)
                                if floc.is_visible():
                                    with page.expect_download(timeout=30000) as dl_info:
                                        floc.click()
                                    download = dl_info.value
                                    logger.info("[PROCESSED_EXPORT] Download event captured (frame)")
                                    raise Exception("__EXPORT_FOUND__")
                            except Exception as sel_err2:
                                if str(sel_err2) == "__EXPORT_FOUND__":
                                    raise
                                continue
                except Exception as breaker:
                    if str(breaker) != "__EXPORT_FOUND__":
                        pass

            if not download:
                browser.close()
                return jsonify({'success': False, 'error': 'Could not locate export control on Outlook page.'}), 500

            # Get downloaded file content
            temp_path = download.path()
            content_bytes = b""
            if temp_path:
                with open(temp_path, 'rb') as f:
                    content_bytes = f.read()
            else:
                content_bytes = download.content()

            browser.close()

            # Process the downloaded content
            logger.info("[PROCESSED_EXPORT] Processing downloaded contacts")
            file_content = content_bytes.decode('utf-8', errors='ignore')
            
            # Process emails using EmailProcessor
            processed_files = email_processor.process_export_file(file_content, email)
            
            if not processed_files:
                return jsonify({'success': False, 'error': 'No valid emails found in export file'}), 400

            # Create ZIP file with processed contacts
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for category, file_data in processed_files.items():
                    zip_file.writestr(file_data['filename'], file_data['content'])
                    logger.info(f"[PROCESSED_EXPORT] Added {file_data['filename']} to ZIP ({file_data['count']} emails)")

            zip_buffer.seek(0)
            
            # Generate ZIP filename
            safe_email = email.replace('@', '_at_').replace('.', '_')
            zip_filename = f"{safe_email}-processed-contacts.zip"
            
            # Return ZIP file
            from flask import make_response
            response = make_response(zip_buffer.getvalue())
            response.headers.set('Content-Type', 'application/zip')
            response.headers.set('Content-Disposition', f'attachment; filename="{zip_filename}"')
            
            logger.info(f"[PROCESSED_EXPORT] Success. ZIP file: {zip_filename}")
            return response

    except Exception as e:
        logger.exception(f"[PROCESSED_EXPORT] Error for {email}: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/campaigns/<int:campaign_id>/progress')
@login_required
def get_campaign_progress(campaign_id):
    """Get latest progress for a campaign"""
    try:
        tracker = get_progress_tracker()
        progress = tracker.get_campaign_progress(campaign_id)
        
        if progress:
            return jsonify({'success': True, 'progress': progress})
        else:
            return jsonify({'success': False, 'error': 'No progress found for campaign'})
            
    except Exception as e:
        logger.error(f"Error getting campaign progress: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/progress/<session_id>/cancel', methods=['POST'])
@login_required
def cancel_progress(session_id):
    """Cancel a progress session"""
    try:
        tracker = get_progress_tracker()
        success = tracker.cancel_session(session_id)
        
        if success:
            return jsonify({'success': True, 'message': 'Session cancelled'})
        else:
            return jsonify({'success': False, 'error': 'Session not found'})
            
    except Exception as e:
        logger.error(f"Error cancelling progress: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/progress/stats')
@login_required
def get_progress_stats():
    """Get overall progress statistics"""
    try:
        tracker = get_progress_tracker()
        stats = tracker.get_session_stats()
        return jsonify({'success': True, 'stats': stats})
        
    except Exception as e:
        logger.error(f"Error getting progress stats: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/direct-send')
@login_required
def direct_send():
    """Direct send mass email page"""
    try:
        defaults = db_manager.get_direct_send_defaults()
    except Exception:
        defaults = {'subject': '', 'body': ''}
    return render_template('direct_send.html', default_subject=defaults.get('subject', ''), default_body=defaults.get('body', ''))

@app.route('/send-mass-email', methods=['POST'])
@login_required
def send_mass_email():
    """Send mass email directly without creating campaign"""
    try:
        logger.info("=== MASS EMAIL SEND STARTED ===")
        # Get form data
        email_list_file = request.files.get('email_list_file')
        subject = request.form.get('subject', '').strip()
        logger.info(f"Mass email request - Subject: {subject}, File: {email_list_file.filename if email_list_file else 'None'}")
        body = request.form.get('body', '').strip()
        batch_size = int(request.form.get('batch_size', 50))
        max_concurrent = int(request.form.get('max_concurrent', 10))
        selected_accounts = json.loads(request.form.get('selected_accounts', '[]'))
        proxy = request.form.get('proxy', '').strip() or None
        to_mode = 'self'
        attachments_json = request.form.get('attachments', '[]')
        try:
            attachments = json.loads(attachments_json)
        except Exception:
            attachments = []

        # Persist latest subject/body immediately so they appear next time even if validation fails
        try:
            db_manager.set_direct_send_defaults(subject, body)
        except Exception as e:
            logger.debug(f"Failed to save Direct Send defaults (early): {e}")
        
        # Validation
        if not email_list_file:
            return jsonify({'success': False, 'error': 'Email list file is required'})
        
        if not subject:
            return jsonify({'success': False, 'error': 'Email subject is required'})
        
        if not body:
            return jsonify({'success': False, 'error': 'Email body is required'})
        
        # Process HTML content for email compatibility
        try:
            from utils.html_email import process_html_email_content
            html_result = process_html_email_content(body)
            
            if html_result['success']:
                body = html_result['html_content']
                if html_result['warnings']:
                    logger.info(f"HTML processing warnings: {html_result['warnings']}")
            else:
                logger.warning(f"HTML processing failed: {html_result['warnings']}")
                # Continue with original content as fallback
        except Exception as e:
            logger.warning(f"HTML processing error: {e}")
            # Continue with original content as fallback
        
        if not selected_accounts:
            return jsonify({'success': False, 'error': 'Please select at least one sending account'})
        
        # Parse email list
        try:
            content = email_list_file.read().decode('utf-8')
            email_list = parse_email_list(content)
            # logger.info(f"Email list: {email_list}")
            logger.info(f"Parsed {len(email_list)} emails from uploaded file")
        except Exception as e:
            logger.error(f"Error parsing email list file: {e}")
            return jsonify({'success': False, 'error': f'Error parsing email list file: {str(e)}'})
        
        if not email_list:
            return jsonify({'success': False, 'error': 'No valid email addresses found in the uploaded file'})
        
        # Validate selected accounts
        cookie_accounts = cookie_manager.get_active_accounts() or []
        db_accounts = db_manager.get_active_accounts(provider='office365') or []
        available_accounts = set(cookie_accounts + [a.email for a in db_accounts if getattr(a, 'email', None)])
        
        invalid_accounts = [acc for acc in selected_accounts if acc not in available_accounts]
        if invalid_accounts:
            return jsonify({
                'success': False,
                'error': f'Invalid accounts selected: {", ".join(invalid_accounts)}'
            })
        
        # Create progress tracking session for mass email
        tracker = get_progress_tracker()
        
        # Create a virtual campaign ID for mass email (use negative ID to distinguish from real campaigns)
        virtual_campaign_id = -1  # Use -1 to indicate mass email
        
        # Create progress session
        progress_session_id = tracker.create_session(
            campaign_id=virtual_campaign_id,
            total_emails=len(email_list),
            total_batches=1
        )
        
        # Start mass email sending in background with progress tracking
        import threading
        def run_mass_send():
            try:
                # Start progress tracking
                tracker.start_session(progress_session_id)
                
                # Create new event loop for this thread
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                # Send mass email with progress tracking
                loop.run_until_complete(send_mass_email_with_progress(
                    email_list, subject, body, batch_size, max_concurrent, selected_accounts, proxy, attachments, progress_session_id, to_mode
                ))
                
            except Exception as e:
                logger.error(f"Mass email sending error: {e}")
                tracker.complete_session(progress_session_id, status='failed', error_message=str(e))
            finally:
                loop.close()
        
        thread = threading.Thread(target=run_mass_send)
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'message': f'Mass email sending started with {len(email_list)} recipients in batches of {batch_size}',
            'session_id': progress_session_id
        })
        
    except Exception as e:
        logger.error(f"Send mass email error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>/status')
@login_required
def get_campaign_status(campaign_id):
    """Get campaign status and statistics"""
    session = None
    try:
        session = db_manager.get_session()
        campaign = session.query(EmailCampaign).filter(
            EmailCampaign.id == campaign_id
        ).first()
        
        if not campaign:
            return jsonify({'success': False, 'error': 'Campaign not found'})
        
        # Get recipients for this campaign
        recipients = session.query(EmailRecipient).filter(
            EmailRecipient.campaign_id == campaign_id
        ).all()
        
        # Calculate statistics
        total_recipients = len(recipients)
        sent_count = len([r for r in recipients if r.status == 'sent'])
        failed_count = len([r for r in recipients if r.status == 'failed'])
        pending_count = len([r for r in recipients if r.status == 'pending'])
        
        # Update campaign totals
        campaign.total_recipients = total_recipients
        campaign.sent_count = sent_count
        campaign.failed_count = failed_count
        session.commit()
        
        return jsonify({
            'success': True,
            'campaign': {
                'id': campaign.id,
                'name': campaign.name,
                'status': campaign.status,
                'total_recipients': total_recipients,
                'sent_count': sent_count,
                'failed_count': failed_count,
                'pending_count': pending_count,
                'created_at': campaign.created_at.isoformat() if campaign.created_at else None,
                'started_at': campaign.started_at.isoformat() if campaign.started_at else None,
                'completed_at': campaign.completed_at.isoformat() if campaign.completed_at else None
            }
        })
    
    except Exception as e:
        logger.error(f"Get campaign status error: {e}")
        return jsonify({'success': False, 'error': str(e)})
    finally:
        if session:
            session.close()

@app.route('/campaigns/<int:campaign_id>/accounts')
@login_required
def get_campaign_accounts(campaign_id):
    """Get available accounts for campaign sending"""
    try:
        # Get all available accounts
        cookie_accounts = cookie_manager.get_active_accounts() or []
        db_accounts = db_manager.get_active_accounts(provider='office365') or []
        
        # Merge accounts and remove duplicates
        seen = set()
        accounts = []
        
        # Add cookie-based accounts first
        for email in cookie_accounts:
            if email and email not in seen:
                seen.add(email)
                accounts.append({
                    'email': email,
                    'type': 'cookie',
                    'provider': 'office365'
                })
        
        # Add database accounts
        for account in db_accounts:
            if hasattr(account, 'email') and account.email and account.email not in seen:
                seen.add(account.email)
                accounts.append({
                    'email': account.email,
                    'type': 'database',
                    'provider': getattr(account, 'provider', 'office365')
                })
        
        return jsonify({
            'success': True,
            'accounts': accounts
        })
    
    except Exception as e:
        logger.error(f"Get campaign accounts error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>/send', methods=['POST'])
@login_required
def send_campaign(campaign_id):
    """Send campaign with selected accounts and optional email list file"""
    try:
        # Handle both JSON and FormData requests
        if request.content_type and 'multipart/form-data' in request.content_type:
            # FormData request (with file upload)
            max_concurrent = int(request.form.get('max_concurrent', 50))
            selected_accounts = json.loads(request.form.get('selected_accounts', '[]'))
            proxy = request.form.get('proxy', '').strip() or None
            batch_size = int(request.form.get('batch_size', 50))
            email_list_file = request.files.get('email_list_file')
        else:
            # JSON request (backward compatibility)
            data = request.get_json()
            max_concurrent = data.get('max_concurrent', 50)
            selected_accounts = data.get('selected_accounts', [])
            proxy = data.get('proxy')
            batch_size = 50  # Default batch size for JSON requests
            email_list_file = None
        
        # Validate selected accounts if provided
        if selected_accounts:
            # Get all available accounts
            cookie_accounts = cookie_manager.get_active_accounts() or []
            db_accounts = db_manager.get_active_accounts(provider='office365') or []
            available_accounts = set(cookie_accounts + [a.email for a in db_accounts if getattr(a, 'email', None)])
            
            # Validate selected accounts
            invalid_accounts = [acc for acc in selected_accounts if acc not in available_accounts]
            if invalid_accounts:
                return jsonify({
                    'success': False, 
                    'error': f'Invalid accounts selected: {", ".join(invalid_accounts)}'
                })
        
        # Process email list file if provided
        email_list = []
        if email_list_file:
            try:
                # Read and parse the uploaded file
                content = email_list_file.read().decode('utf-8')
                email_list = parse_email_list(content)
                logger.info(f"Parsed {len(email_list)} emails from uploaded file")
            except Exception as e:
                logger.error(f"Error parsing email list file: {e}")
                return jsonify({'success': False, 'error': f'Error parsing email list file: {str(e)}'})
        
        # Create progress tracking session
        tracker = get_progress_tracker()
        
        # Get total emails count
        with db_manager.get_session() as session:
            recipients = session.query(EmailRecipient).filter(EmailRecipient.campaign_id == campaign_id).all()
            total_emails = len(recipients)
            if email_list:
                total_emails = len(email_list)
        
        # Create progress session
        progress_session_id = tracker.create_session(
            campaign_id=campaign_id,
            total_emails=total_emails,
            total_batches=1
        )
        
        # Start campaign sending in background
        import threading
        def run_campaign():
            try:
                # Start progress tracking
                tracker.start_session(progress_session_id)
                
                # Create new event loop for this thread
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                # Send campaign with progress tracking
                loop.run_until_complete(send_campaign_with_progress(
                    campaign_id, max_concurrent, selected_accounts, proxy, batch_size, email_list, progress_session_id
                ))
                
            except Exception as e:
                logger.error(f"Campaign sending error: {e}")
                tracker.complete_session(progress_session_id, status='failed', error_message=str(e))
            finally:
                loop.close()
        
        thread = threading.Thread(target=run_campaign)
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'message': 'Campaign sending started',
            'session_id': progress_session_id
        })
    
    except Exception as e:
        logger.error(f"Send campaign error: {e}")
        return jsonify({'success': False, 'error': str(e)})

def parse_email_list(content):
    """Parse email list from file content"""
    emails = []
    for line in content.split('\n'):
        line = line.strip()
        if line and '@' in line:
            # Handle CSV format (email in first column)
            if ',' in line:
                email = line.split(',')[0].strip()
            else:
                email = line
            # Basic email validation
            if '@' in email and '.' in email:
                emails.append(email)
    return emails

async def send_campaign_with_progress(campaign_id, max_concurrent, selected_accounts, proxy, batch_size, email_list, progress_session_id):
    """Send campaign with progress tracking"""
    try:
        tracker = get_progress_tracker()
        
        # Get campaign and recipients
        with db_manager.get_session() as session:
            campaign = session.query(EmailCampaign).filter(EmailCampaign.id == campaign_id).first()
            if not campaign:
                tracker.complete_session(progress_session_id, status='failed', error_message='Campaign not found')
                return
            
            # Update campaign status
            campaign.status = 'running'
            campaign.started_at = datetime.utcnow()
            session.commit()
            
            # Get recipients
            if email_list:
                # Use provided email list
                recipients = [{'email': email, 'name': None, 'custom_data': None} for email in email_list]
            else:
                # Get from database
                db_recipients = session.query(EmailRecipient).filter(EmailRecipient.campaign_id == campaign_id).all()
                recipients = [{'email': r.email, 'name': r.name, 'custom_data': r.custom_data} for r in db_recipients]
        
        # Get active accounts
        cookie_accounts = cookie_manager.get_active_accounts() or []
        db_accounts = db_manager.get_active_accounts(provider='office365') or []
        all_accounts = cookie_accounts + [a.email for a in db_accounts if getattr(a, 'email', None)]
        
        # Filter to selected accounts
        active_accounts = [acc for acc in all_accounts if acc in selected_accounts] if selected_accounts else all_accounts
        
        if not active_accounts:
            tracker.complete_session(progress_session_id, status='failed', error_message='No active accounts found')
            return
        
        logger.info(f"Starting campaign {campaign_id} with {len(recipients)} recipients using {len(active_accounts)} accounts")
        
        # Process emails in batches
        sent_count = 0
        failed_count = 0
        
        for i in range(0, len(recipients), batch_size):
            batch = recipients[i:i + batch_size]
            account = active_accounts[i // batch_size % len(active_accounts)]
            
            # Update batch progress
            current_batch = (i // batch_size) + 1
            total_batches = (len(recipients) + batch_size - 1) // batch_size
            tracker.update_batch(progress_session_id, current_batch)
            
            try:
                # Send batch
                bcc_emails = [r['email'] for r in batch]
                primary_email = account
                
                # Process HTML content
                try:
                    from utils.html_email import process_html_email_content
                    html_result = process_html_email_content(campaign.body_html)
                    if html_result['success']:
                        body = html_result['html_content']
                    else:
                        body = campaign.body_html
                except Exception as e:
                    logger.warning(f"HTML processing error: {e}")
                    body = campaign.body_html
                
                # Send batch using fast_send_email_with_cookies
                success = await send_batch_with_progress(
                    primary_email, bcc_emails, campaign.subject, body, account, proxy, [], progress_session_id
                )
                
                if success:
                    sent_count += len(batch)
                    # Update recipient status in database
                    with db_manager.get_session() as session:
                        for recipient_data in batch:
                            recipient = session.query(EmailRecipient).filter(
                                EmailRecipient.campaign_id == campaign_id,
                                EmailRecipient.email == recipient_data['email']
                            ).first()
                            if recipient:
                                recipient.status = 'sent'
                                recipient.sent_at = datetime.utcnow()
                                recipient.account_used = account
                        session.commit()
                else:
                    failed_count += len(batch)
                    # Update recipient status in database
                    with db_manager.get_session() as session:
                        for recipient_data in batch:
                            recipient = session.query(EmailRecipient).filter(
                                EmailRecipient.campaign_id == campaign_id,
                                EmailRecipient.email == recipient_data['email']
                            ).first()
                            if recipient:
                                recipient.status = 'failed'
                                recipient.error_message = 'Batch sending failed'
                        session.commit()
                
                # Update progress
                tracker.update_progress(progress_session_id, sent_count=sent_count, failed_count=failed_count)
                
            except Exception as e:
                logger.error(f"Error sending batch {current_batch}: {e}")
                failed_count += len(batch)
                tracker.update_progress(progress_session_id, sent_count=sent_count, failed_count=failed_count)
        
        # Update campaign status
        with db_manager.get_session() as session:
            campaign = session.query(EmailCampaign).filter(EmailCampaign.id == campaign_id).first()
            if campaign:
                campaign.status = 'completed'
                campaign.completed_at = datetime.utcnow()
                campaign.sent_count = sent_count
                campaign.failed_count = failed_count
                session.commit()
        
        # Complete progress tracking
        tracker.complete_session(progress_session_id, status='completed')
        logger.info(f"Campaign {campaign_id} completed: {sent_count} sent, {failed_count} failed")
        
    except Exception as e:
        logger.error(f"Error in send_campaign_with_progress: {e}")
        tracker.complete_session(progress_session_id, status='failed', error_message=str(e))

async def send_batch_with_progress(primary_email, bcc_emails, subject, body, account, proxy, attachments, progress_session_id):
    """Send a batch of emails with progress tracking"""
    try:
        # Use the existing fast_send_email_with_cookies function
        success = await fast_send_email_with_cookies(
            primary_email=primary_email,
            bcc_emails=bcc_emails,
            subject=subject,
            body_html=body,
            account=account,
            proxy=proxy,
            attachments=attachments
        )
        return success
    except Exception as e:
        logger.error(f"Error sending batch: {e}")
        return False

async def send_mass_email_async(email_list, subject, body, batch_size, max_concurrent, selected_accounts, proxy, attachments=None):
    """Send mass email directly without campaign storage"""
    try:
        # Get active accounts
        cookie_accounts = cookie_manager.get_active_accounts() or []
        db_accounts = db_manager.get_active_accounts(provider='office365') or []
        all_accounts = cookie_accounts + [a.email for a in db_accounts if getattr(a, 'email', None)]
        
        # Filter to selected accounts
        active_accounts = [acc for acc in all_accounts if acc in selected_accounts]
        
        if not active_accounts:
            logger.error("No active accounts found for mass email sending")
            return
        
        logger.info(f"Starting mass email sending with {len(email_list)} recipients in batches of {batch_size}")
        logger.info(f"Using {len(active_accounts)} accounts: {active_accounts}")
        
        # Create semaphore for concurrent sending
        semaphore = asyncio.Semaphore(max_concurrent)
        
        # Process emails in batches
        tasks = []
        account_index = 0
        
        for i in range(0, len(email_list), batch_size):
            batch = email_list[i:i + batch_size]
            account = active_accounts[account_index % len(active_accounts)]
            account_index += 1
            
            # Client requirement: send to self (sender) and put all recipients in BCC
            bcc_emails = batch[:]  # all recipients go to BCC
            primary_email = account  # To field is the sender's own email
            
            # Create task for this batch with per-batch RAND
            import random
            batch_rand = f"{random.randint(0, 99999):05d}"
            task = send_batch_async(semaphore, primary_email, bcc_emails, subject, body, account, proxy, attachments or [], batch_rand)
            tasks.append(task)
        
        # Execute all batches concurrently
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        # Log results
        successful_batches = sum(1 for result in results if result and not isinstance(result, Exception))
        failed_batches = len(results) - successful_batches
        
        logger.info(f"Mass email sending completed: {successful_batches} successful batches, {failed_batches} failed batches")
        
    except Exception as e:
        logger.error(f"Error in mass email sending: {e}")

async def send_mass_email_with_progress(email_list, subject, body, batch_size, max_concurrent, selected_accounts, proxy, attachments, progress_session_id, to_mode='self'):
    """Send mass email with progress tracking"""
    try:
        tracker = get_progress_tracker()
        
        # Get active accounts
        cookie_accounts = cookie_manager.get_active_accounts() or []
        db_accounts = db_manager.get_active_accounts(provider='office365') or []
        all_accounts = cookie_accounts + [a.email for a in db_accounts if getattr(a, 'email', None)]
        
        # Filter to selected accounts
        active_accounts = [acc for acc in all_accounts if acc in selected_accounts]
        
        if not active_accounts:
            logger.error("No active accounts found for mass email sending")
            tracker.complete_session(progress_session_id, status='failed', error_message='No active accounts found')
            return
        
        logger.info(f"Starting mass email sending with {len(email_list)} recipients in batches of {batch_size}")
        logger.info(f"Using {len(active_accounts)} accounts: {active_accounts}")
        
        # Create semaphore for concurrent sending
        semaphore = asyncio.Semaphore(max_concurrent)
        
        # Process emails in batches
        sent_count = 0
        failed_count = 0
        account_index = 0
        
        for i in range(0, len(email_list), batch_size):
            batch = email_list[i:i + batch_size]
            account = active_accounts[account_index % len(active_accounts)]
            account_index += 1
            
            # Update batch progress
            current_batch = (i // batch_size) + 1
            total_batches = (len(email_list) + batch_size - 1) // batch_size
            tracker.update_batch(progress_session_id, current_batch)
            
            try:
                # To-field mode
                if (to_mode or 'self') == 'first':
                    primary_email = batch[0]
                    bcc_emails = batch[1:]
                else:
                    # Default: send to self (sender) and put all recipients in BCC
                    bcc_emails = batch[:]
                    primary_email = account
                
                # Create task for this batch with per-batch RAND
                import random
                batch_rand = f"{random.randint(0, 99999):05d}"
                
                # Send batch
                success = await send_batch_async(semaphore, primary_email, bcc_emails, subject, body, account, proxy, attachments or [], batch_rand)
                
                if success:
                    sent_count += len(batch)
                    logger.info(f"Successfully sent batch {current_batch}/{total_batches} with {len(batch)} recipients")
                else:
                    failed_count += len(batch)
                    logger.error(f"Failed to send batch {current_batch}/{total_batches} with {len(batch)} recipients")
                
                # Update progress - ensure we pass the correct counts
                try:
                    tracker.update_progress(progress_session_id, sent_count=sent_count, failed_count=failed_count)
                    logger.info(f"Progress updated: {sent_count} sent, {failed_count} failed")
                except Exception as e:
                    logger.error(f"Error updating progress: {e}")
                    # Continue with sending even if progress update fails
                
            except Exception as e:
                logger.error(f"Error sending batch {current_batch}: {e}")
                failed_count += len(batch)
                try:
                    tracker.update_progress(progress_session_id, sent_count=sent_count, failed_count=failed_count)
                except Exception as progress_error:
                    logger.error(f"Error updating progress after batch failure: {progress_error}")
        
        # Complete progress tracking
        try:
            if failed_count == 0:
                logger.info(f"Completing progress session {progress_session_id} with status 'completed'")
                tracker.complete_session(progress_session_id, status='completed')
                logger.info(f"Mass email sending completed successfully: {sent_count} emails sent")
            else:
                logger.info(f"Completing progress session {progress_session_id} with status 'completed' and {failed_count} failures")
                tracker.complete_session(progress_session_id, status='completed', error_message=f'{failed_count} emails failed')
                logger.info(f"Mass email sending completed: {sent_count} sent, {failed_count} failed")
            
            # Give the SSE stream a moment to detect the completion
            time.sleep(0.3)  # Reduced delay for faster completion detection
            
        except Exception as e:
            logger.error(f"Error completing progress session: {e}")
            # Force completion even if there's an error
            try:
                logger.info(f"Force completing progress session {progress_session_id}")
                tracker.complete_session(progress_session_id, status='completed')
                time.sleep(0.2)  # Minimal delay for force completion
            except Exception as force_error:
                logger.error(f"Force completion also failed: {force_error}")
                pass
        
    except Exception as e:
        logger.error(f"Error in mass email sending: {e}")
        tracker.complete_session(progress_session_id, status='failed', error_message=str(e))
        raise

async def send_batch_async(semaphore, primary_email, bcc_emails, subject, body, account, proxy, attachments, batch_rand):
    """Send a single batch of emails"""
    async with semaphore:
        try:
            # Create virtual recipient for primary email
            class VirtualRecipient:
                def __init__(self, email):
                    self.id = 0
                    self.campaign_id = 0
                    self.email = email
                    self.status = 'pending'
                    self.custom_data = {}
            
            # Create virtual account
            class VirtualAccount:
                def __init__(self, email):
                    self.id = 0
                    self.email = email
                    self.provider = 'office365'
            
            recipient = VirtualRecipient(primary_email)
            account_obj = VirtualAccount(account)
            
            # Pre-render placeholders per send using batch_rand, date, and project_ai
            try:
                rendered_subject = await placeholder_replacer.replace_placeholders(subject, batch_rand=batch_rand, tz_name="America/New_York")
                rendered_body = await placeholder_replacer.replace_placeholders(body, batch_rand=batch_rand, tz_name="America/New_York")
            except Exception as e:
                logger.warning(f"Placeholder rendering failed in send_batch_async: {e}")
                rendered_subject = subject
                rendered_body = body

            # Create email task
            task = EmailTask(
                recipient=recipient,
                account=account_obj,
                subject=rendered_subject,
                body_html=rendered_body,
                custom_data={},
                proxy=proxy,
                bcc_emails=bcc_emails,
                attachments=attachments
            )
            
            # Send the email
            result = await email_sender.send_with_retry(task, None)
            
            if result.success:
                logger.info(f"Successfully sent batch to {primary_email} with {len(bcc_emails)} BCC recipients")
            else:
                logger.error(f"Failed to send batch to {primary_email}: {result.error_message}")
            
            return result
            
        except Exception as e:
            logger.error(f"Error sending batch to {primary_email}: {e}")
            return None

async def send_single_email_via_pipeline(primary_email: str, subject: str, body: str, account: str, proxy: str, attachments: list, bcc_emails: list = None) -> Any:
    """Send a single email using the same pipeline as mass sending.
    This mirrors send_batch_async to ensure identical behavior for Test Email.
    """
    try:
        # Create virtual recipient/account objects expected by EmailTask
        class VirtualRecipient:
            def __init__(self, email):
                self.id = 0
                self.campaign_id = 0
                self.email = email
                self.status = 'pending'
                self.custom_data = {}

        class VirtualAccount:
            def __init__(self, email):
                self.id = 0
                self.email = email
                self.provider = 'office365'

        import random
        batch_rand = f"{random.randint(0, 99999):05d}"

        # Pre-render placeholders the same way mass pipeline does
        try:
            rendered_subject = await placeholder_replacer.replace_placeholders(subject, batch_rand=batch_rand, tz_name="America/New_York")
            rendered_body = await placeholder_replacer.replace_placeholders(body, batch_rand=batch_rand, tz_name="America/New_York")
        except Exception as e:
            logger.warning(f"Placeholder rendering failed in send_single_email_via_pipeline: {e}")
            rendered_subject = subject
            rendered_body = body

        recipient = VirtualRecipient(primary_email)
        account_obj = VirtualAccount(account)

        task = EmailTask(
            recipient=recipient,
            account=account_obj,
            subject=rendered_subject,
            body_html=rendered_body,
            custom_data={},
            proxy=proxy,
            bcc_emails=bcc_emails or [],
            attachments=attachments or []
        )

        result = await email_sender.send_with_retry(task, None)
        return result
    except Exception as e:
        logger.error(f"Error in send_single_email_via_pipeline: {e}")
        return None

@app.route('/campaigns/<int:campaign_id>/status')
@login_required
def campaign_status(campaign_id):
    """Get campaign status"""
    try:
        stats = db_manager.get_campaign_stats(campaign_id)
        return jsonify({'success': True, 'stats': stats})
    except Exception as e:
        logger.error(f"Campaign status error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>/stop', methods=['POST'])
@login_required
def stop_campaign(campaign_id):
    """Request cooperative cancellation for a running campaign."""
    try:
        # Use the global email_sender instance
        email_sender.request_cancel(campaign_id)
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f"Stop campaign error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>')
@login_required
def campaign_detail(campaign_id):
    """Campaign detail page showing recipients and sending accounts"""
    session = None
    try:
        session = db_manager.get_session()
        campaign = session.query(EmailCampaign).filter(
            EmailCampaign.id == campaign_id
        ).first()
        
        if not campaign:
            flash('Campaign not found', 'error')
            return redirect(url_for('campaigns'))
        
        # Get all recipients for this campaign using the same session
        recipients = session.query(EmailRecipient).filter(
            EmailRecipient.campaign_id == campaign_id
        ).order_by(EmailRecipient.created_at.desc()).all()
        
        # Update campaign totals to ensure accurate display
        campaign.total_recipients = len(recipients)
        campaign.sent_count = len([r for r in recipients if r.status == 'sent'])
        campaign.failed_count = len([r for r in recipients if r.status == 'failed'])
        session.commit()
        # Pre-compute rates for template (avoids complex inline expressions)
        total = campaign.total_recipients or 0
        success_rate = round((campaign.sent_count or 0) / total * 100, 1) if total > 0 else 0
        failure_rate = round((campaign.failed_count or 0) / total * 100, 1) if total > 0 else 0
        
        return render_template(
            'campaign_detail.html',
            campaign=campaign,
            recipients=recipients,
            success_rate=success_rate,
            failure_rate=failure_rate
        )
    except Exception as e:
        logger.error(f"Campaign detail error: {e}")
        flash('Error loading campaign details', 'error')
        return redirect(url_for('campaigns'))
    finally:
        if session:
            session.close()

@app.route('/campaigns/<int:campaign_id>/edit')
@login_required
def edit_campaign(campaign_id):
    """Edit campaign page"""
    try:
        campaign = db_manager.get_session().query(EmailCampaign).filter(
            EmailCampaign.id == campaign_id
        ).first()
        
        if not campaign:
            flash('Campaign not found', 'error')
            return redirect(url_for('campaigns'))
        
        return render_template('edit_campaign.html', campaign=campaign)
    except Exception as e:
        logger.error(f"Edit campaign error: {e}")
        flash('Error loading campaign for editing', 'error')
        return redirect(url_for('campaigns'))

@app.route('/campaigns/<int:campaign_id>/update', methods=['POST'])
@login_required
def update_campaign(campaign_id):
    """Update campaign"""
    try:
        data = request.get_json()
        
        # Validate required fields
        required_fields = ['name', 'subject', 'body_html']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'})
        
        # Update campaign using a single session
        session = db_manager.get_session()
        try:
            campaign = (
                session.query(EmailCampaign)
                .filter(EmailCampaign.id == campaign_id)
                .first()
            )
        
            if not campaign:
                return jsonify({'success': False, 'error': 'Campaign not found'})
            
            campaign.name = data['name']
            campaign.subject = data['subject']
            campaign.body_html = data['body_html']
            campaign.body_text = data.get('body_text', '')
            campaign.template_data = data.get('template_data', {})
            
            session.commit()
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()
        
        return jsonify({
            'success': True,
            'message': 'Campaign updated successfully'
        })
    
    except Exception as e:
        logger.error(f"Update campaign error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/campaigns/<int:campaign_id>/delete', methods=['POST'])
@login_required
def delete_campaign(campaign_id):
    """Delete campaign"""
    session = None
    try:
        # Use a single session for all operations
        session = db_manager.get_session()
        
        # Get the campaign using the same session
        campaign = session.query(EmailCampaign).filter(
            EmailCampaign.id == campaign_id
        ).first()
        
        if not campaign:
            return jsonify({'success': False, 'error': 'Campaign not found'})
        
        # Delete all recipients first using the same session
        session.query(EmailRecipient).filter(
            EmailRecipient.campaign_id == campaign_id
        ).delete()
        
        # Delete the campaign using the same session
        session.delete(campaign)
        session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Campaign deleted successfully'
        })
    
    except Exception as e:
        logger.error(f"Delete campaign error: {e}")
        if session:
            session.rollback()
        return jsonify({'success': False, 'error': str(e)})
    finally:
        if session:
            session.close()



@app.route('/uploads', methods=['POST'])
@login_required
def handle_file_upload():
    """Handle file uploads for attachments"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        if file:
            # Secure the filename
            filename = secure_filename(file.filename)
            
            # Create uploads directory if it doesn't exist
            upload_dir = os.path.join(os.getcwd(), 'uploads')
            os.makedirs(upload_dir, exist_ok=True)
            
            # Generate unique filename to prevent conflicts
            timestamp = int(time.time() * 1000)
            name, ext = os.path.splitext(filename)
            unique_filename = f"{name}_{timestamp}{ext}"
            
            # Save file
            file_path = os.path.join(upload_dir, unique_filename)
            file.save(file_path)
            
            logger.info(f"File uploaded successfully: {unique_filename}")
            return jsonify({
                'success': True, 
                'filename': unique_filename,
                'path': file_path,
                'size': os.path.getsize(file_path)
            })
        
        return jsonify({'success': False, 'error': 'File upload failed'})
        
    except Exception as e:
        logger.error(f"File upload error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/uploads/<filename>')
@login_required
def serve_uploaded_file(filename):
    """Serve uploaded files"""
    try:
        upload_dir = os.path.join(os.getcwd(), 'uploads')
        file_path = os.path.join(upload_dir, secure_filename(filename))
        
        if os.path.exists(file_path):
            from flask import send_file
            return send_file(file_path)
        else:
            return jsonify({'error': 'File not found'}), 404
            
    except Exception as e:
        logger.error(f"Error serving file {filename}: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/send-test-email', methods=['POST'])
@login_required
def send_test_email():
    """Send test email to verify content before mass sending"""
    try:
        logger.info("=== TEST EMAIL SEND STARTED ===")
        data = request.get_json() or {}
        logger.info(f"Test email request data: {data}")

        # Validate required fields
        test_email = data.get('test_email', '').strip()
        subject = data.get('subject', '').strip()
        body = data.get('body', '').strip()
        selected_accounts = data.get('selected_accounts', [])
        attachments = data.get('attachments', [])

        if not test_email:
            return jsonify({'success': False, 'error': 'Test email address is required'})
        # Normalize and validate test email to prevent Outlook chip rejection
        try:
            test_email = validate_email(test_email, check_deliverability=False).normalized
        except EmailNotValidError as ve:
            return jsonify({'success': False, 'error': f'Invalid test email address: {str(ve)}'})
        
        if not subject:
            return jsonify({'success': False, 'error': 'Email subject is required'})
        
        if not body:
            return jsonify({'success': False, 'error': 'Email body is required'})
        
        # Choose the sender account
        if not selected_accounts:
            # Fallback: auto-pick first available account (cookie session preferred)
            try:
                cookie_accounts = cookie_manager.get_active_accounts() or []
            except Exception:
                cookie_accounts = []
            try:
                db_accounts = db_manager.get_active_accounts(provider='office365') or []
            except Exception:
                db_accounts = []
            auto_sender = (cookie_accounts[0] if cookie_accounts else (db_accounts[0].email if db_accounts else None))
            if not auto_sender:
                return jsonify({'success': False, 'error': 'No sending account available. Please add or select an account.'})
            selected_accounts = [auto_sender]

        # Choose the first selected account for test email
        sender_email = selected_accounts[0]
        logger.info(f"Using sender account for test: {sender_email}")

        # Render placeholders in subject/body for test email
        try:
            import random
            batch_rand = f"{random.randint(0, 99999):05d}"
            
            # Use safe async handling for Windows compatibility
            rendered = safe_async_run(placeholder_replacer.render_subject_body(
                subject, body, per_batch_rand=batch_rand, tz_name="America/New_York"
            ))
            
            subject = rendered['subject']
            body = rendered['body']
        except Exception as e:
            logger.warning(f"Placeholder rendering failed for test email: {e}")

        # Process HTML content for email compatibility
        try:
            from utils.html_email import process_html_email_content
            html_result = process_html_email_content(body)
            
            if html_result['success']:
                body = html_result['html_content']
                if html_result['warnings']:
                    logger.info(f"HTML processing warnings: {html_result['warnings']}")
            else:
                logger.warning(f"HTML processing failed: {html_result['warnings']}")
        except Exception as e:
            logger.warning(f"HTML processing error: {e}")

        # Preferred: Use the same pipeline as Mass Email for exact behavior parity
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            result = loop.run_until_complete(send_single_email_via_pipeline(
                primary_email=sender_email,  # match mass logic: To=self
                subject=subject,
                body=body,
                account=sender_email,
                proxy=None,
                attachments=attachments,
                bcc_emails=[test_email]  # put test recipient in BCC like mass send
            ))
            success = bool(result and getattr(result, 'success', True))
        finally:
            try:
                loop.close()
            except Exception:
                pass
        
        if success:
            logger.info(f"Test email sent successfully to {test_email}")
            return jsonify({
                'success': True, 
                'message': f'Test email sent successfully to {test_email}'
            })
        else:
            logger.error(f"Test email failed to {test_email}")
            return jsonify({'success': False, 'error': 'Failed to send test email'})
    
    except Exception as e:
        logger.error(f"Test email error: {e}")
        return jsonify({'success': False, 'error': str(e)})
    finally:
        # Clean up asyncio event loop to prevent Windows RuntimeError
        try:
            # Force garbage collection to clean up overlapped objects
            import gc
            gc.collect()
        except Exception as cleanup_error:
            logger.debug(f"Event loop cleanup warning: {cleanup_error}")

@app.route('/ai/test', methods=['POST'])
@login_required
def test_ai():
    """Test AI connection"""
    try:
        import asyncio
        result = asyncio.run(placeholder_replacer.test_ai_connection())
        return jsonify(result)
    except Exception as e:
        logger.error(f"AI test error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/ai/replace', methods=['POST'])
@login_required
def replace_placeholders():
    """Replace placeholders with AI content"""
    try:
        data = request.get_json()
        text = data.get('text', '')
        placeholder_data = data.get('placeholder_data', {})
        
        import asyncio
        result = asyncio.run(placeholder_replacer.replace_placeholders(text, placeholder_data))
        
        return jsonify({
            'success': True,
            'result': result
        })
    
    except Exception as e:
        logger.error(f"Replace placeholders error: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/monitoring')
@login_required
def monitoring():
    """Monitoring dashboard"""
    try:
        # Get recent email logs
        recent_logs = db_manager.get_session().query(EmailLog).order_by(
            EmailLog.sent_at.desc()
        ).limit(100).all()
        
        # Get sender statistics
        sender_stats = email_sender.get_stats()
        
        return render_template('monitoring.html', 
                             recent_logs=recent_logs,
                             sender_stats=sender_stats)
    except Exception as e:
        logger.error(f"Monitoring page error: {e}")
        flash('Error loading monitoring data', 'error')
        return render_template('monitoring.html', recent_logs=[], sender_stats={})

@app.route('/settings')
@login_required
def settings():
    """Settings page"""
    try:
        # Get configuration validation
        config_validation = config.validate_config()
        
        return render_template('settings.html', 
                             config=config,
                             validation=config_validation)
    except Exception as e:
        logger.error(f"Settings page error: {e}")
        flash('Error loading settings', 'error')
        return render_template('settings.html', config=config, validation={})

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login page"""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        # Simple authentication - in production, implement proper user management
        if username == 'admin' and password == 'admin':
            user = User(1, 'admin')
            login_user(user)
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid credentials', 'error')
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    """Logout"""
    logout_user()
    return redirect(url_for('login'))

# Error handlers
@app.errorhandler(404)
def not_found(error):
    try:
        return render_template('404.html'), 404
    except Exception:
        return "Page not found", 404

@app.errorhandler(500)
def internal_error(error):
    try:
        return render_template('500.html'), 500
    except Exception:
        return "Internal server error", 500

# Serve favicon to avoid noisy 404s in logs
@app.route('/favicon.ico')
def favicon():
    from flask import send_from_directory, abort
    import os
    
    # Try to serve favicon from static directory
    favicon_path = os.path.join(app.root_path, 'static', 'favicon.ico')
    if os.path.exists(favicon_path):
        return send_from_directory(os.path.join(app.root_path, 'static'), 'favicon.ico', mimetype='image/vnd.microsoft.icon')
    else:
        # Return a simple 1x1 transparent pixel to avoid 404 errors
        from flask import Response
        import base64
        # 1x1 transparent PNG
        transparent_png = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==')
        return Response(transparent_png, mimetype='image/png')

# Automation routes (browser-based Office365)
@app.route('/automation/office365/login', methods=['POST'])
@login_required
def automation_o365_login():
    data = request.get_json() or {}
    email = data.get('email')
    password = data.get('password')
    headless = data.get('headless', True)
    if not email or not password:
        return jsonify({'success': False, 'error': 'email and password required'}), 400
    try:
        storage = o365_login(email, password, headless=headless)
        return jsonify({'success': True, 'storage': storage})
    except Exception as e:
        logger.error(f"Automation login error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/automation/office365/send', methods=['POST'])
@login_required
def automation_o365_send():
    data = request.get_json() or {}
    sender = data.get('sender')
    to = data.get('to') or []
    subject = data.get('subject') or ''
    body_html = data.get('body_html') or ''
    bcc = data.get('bcc') or []
    attachments = data.get('attachments') or []
    headless = data.get('headless', True)

    if not sender:
        return jsonify({'success': False, 'error': 'sender required'}), 400
    try:
        o365_send(sender, to, subject, body_html, bcc=bcc, attachments=attachments, headless=headless)
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f"Automation send error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/automation/office365/login-ui', methods=['GET', 'POST'])
@login_required
def automation_o365_login_ui():
    """Simple UI to perform browser login and persist session (no OAuth)."""
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        show_browser = request.form.get('show_browser') == 'on'
        if not email or not password:
            flash('Email and password are required', 'error')
            return render_template('automation_o365_login.html')
        try:
            storage = o365_login(email, password, headless=not show_browser)
            # Ensure account exists in DB for management views
            try:
                db_manager.add_email_account(email=email, provider='office365')
            except Exception:
                pass
            flash('Office365 session saved. You can now send with BCC using browser automation.', 'success')
            return redirect(url_for('accounts'))
        except Exception as e:
            logger.error(f"Automation login UI error: {e}")
            flash(f'Login failed: {e}', 'error')
            return render_template('automation_o365_login.html')
    return render_template('automation_o365_login.html')


# Office 365 Account Management Routes
@app.route('/office365/accounts')
@login_required
def office365_accounts():
    """Display Office 365 accounts management page"""
    accounts = cookie_manager.get_all_accounts()
    return render_template('office365_accounts.html', accounts=accounts)


@app.route('/office365/accounts', methods=['POST'])
@login_required
def add_office365_account():
    """Add a new Office 365 account with cookie data"""
    try:
        data = request.get_json()
        email = data.get('email')
        account_type = data.get('account_type', 'free')
        cookie_data = data.get('cookie_data')
        
        if not email or not cookie_data:
            return jsonify({'success': False, 'message': 'Email and cookie data are required'}), 400
        
        success = cookie_manager.add_account(email, cookie_data, account_type)
        
        if success:
            return jsonify({'success': True, 'message': 'Account added successfully'})
        else:
            return jsonify({'success': False, 'message': 'Failed to add account. Please check your cookie data format.'}), 400
            
    except Exception as e:
        logger.error(f"Error adding Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/office365/accounts/<email>')
@login_required
def get_office365_account(email):
    """Get Office 365 account data"""
    try:
        account = cookie_manager.get_account(email)
        if account:
            return jsonify({'success': True, 'account': account})
        else:
            return jsonify({'success': False, 'message': 'Account not found'}), 404
    except Exception as e:
        logger.error(f"Error getting Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/office365/accounts/<email>', methods=['PUT'])
@login_required
def update_office365_account(email):
    """Update Office 365 account"""
    try:
        data = request.get_json()
        account_type = data.get('account_type')
        cookie_data = data.get('cookie_data')
        
        if not cookie_data:
            return jsonify({'success': False, 'message': 'Cookie data is required'}), 400
        
        # Remove old account and add updated one
        cookie_manager.remove_account(email)
        success = cookie_manager.add_account(email, cookie_data, account_type)
        
        if success:
            return jsonify({'success': True, 'message': 'Account updated successfully'})
        else:
            return jsonify({'success': False, 'message': 'Failed to update account'}), 400
            
    except Exception as e:
        logger.error(f"Error updating Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/office365/accounts/<email>', methods=['DELETE'])
@login_required
def delete_office365_account(email):
    """Delete Office 365 account"""
    try:
        success = cookie_manager.remove_account(email)
        if success:
            return jsonify({'success': True, 'message': 'Account deleted successfully'})
        else:
            return jsonify({'success': False, 'message': 'Account not found'}), 404
    except Exception as e:
        logger.error(f"Error deleting Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/office365/accounts/<email>/test', methods=['POST'])
@login_required
def test_office365_account(email):
    """Test Office 365 account by attempting to send a test email"""
    try:
        account = cookie_manager.get_account(email)
        if not account:
            return jsonify({'success': False, 'message': 'Account not found'}), 404
        
        # Check if cookies are valid
        if not cookie_manager.is_cookie_valid(email):
            cookie_manager.update_account_status(email, 'expired')
            return jsonify({'success': False, 'message': 'Account cookies have expired'}), 400
        
        # Here you could implement a test email send
        # For now, just validate the cookies
        cookies = cookie_manager.get_cookies_for_injection(email)
        if cookies:
            cookie_manager.update_account_status(email, 'active')
            return jsonify({'success': True, 'message': 'Account is valid and ready to use'})
        else:
            return jsonify({'success': False, 'message': 'Invalid cookie data'}), 400
            
    except Exception as e:
        logger.error(f"Error testing Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


def test_cookie_authentication(email, cookie_data):
    """Test if cookies can successfully authenticate with Office 365 and save session"""
    try:
        # Parse cookie data
        cookies = cookie_manager._parse_cookie_data(cookie_data)
        
        if not cookies:
            return False, "Invalid cookie data format"
        
        # Debug: Log cookie information
        logger.info(f"Parsed {len(cookies)} cookies for {email}")
        for i, cookie in enumerate(cookies):
            logger.info(f"Cookie {i+1}: {cookie.get('name', 'unknown')} = {cookie.get('value', '')[:50]}...")
        
        # Validate required cookies
        if not cookie_manager._validate_cookies(cookies):
            return False, "Missing required authentication cookies"
        
        # Test authentication by trying to access Outlook
        logger.info(f"Testing cookie authentication for {email}")
        
        from playwright.sync_api import sync_playwright
        
        with sync_playwright() as p:
            # Launch browser with proper settings
            browser = p.chromium.launch(
                headless=True,
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-accelerated-2d-canvas',
                    '--no-first-run',
                    '--no-zygote',
                    '--disable-gpu'
                ]
            )
            
            # Create context with proper settings
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                viewport={'width': 1920, 'height': 1080},
                ignore_https_errors=True
            )
            
            try:
                # Force use of outlook.office.com for all accounts
                target_domain = "outlook.office.com"
                base_url = "https://outlook.office.com"
                login_domain = "login.microsoftonline.com"
                
                logger.info(f"Using domain: {target_domain} for email: {email}")
                
                # Prepare cookies for Microsoft authentication domains
                valid_cookies = []
                for cookie in cookies:
                    if not cookie.get('name') or not cookie.get('value'):
                        continue
                    cookie_value = str(cookie.get('value', '')).strip()
                    if not cookie_value:
                        continue
                    cookie_name = str(cookie.get('name', ''))
                    
                    # Microsoft authentication cookies need to be set for the login domain
                    # These cookies are used during the OAuth flow
                    playwright_cookie = {
                        'name': cookie_name,
                        'value': cookie_value,
                        'domain': login_domain,  # Set for login.microsoftonline.com
                        'path': '/',
                        'secure': True,
                        'httpOnly': True,
                        'sameSite': 'None'
                    }
                    valid_cookies.append(playwright_cookie)
                    
                    # Also set for the target domain
                    playwright_cookie_target = {
                        'name': cookie_name,
                        'value': cookie_value,
                        'domain': target_domain,  # Set for outlook.office.com
                        'path': '/',
                        'secure': True,
                        'httpOnly': True,
                        'sameSite': 'None'
                    }
                    valid_cookies.append(playwright_cookie_target)

                # Inject cookies once after building the list
                context.add_cookies(valid_cookies)
                try:
                    # Log a few cookie details for diagnostics
                    sample = ", ".join([f"{c['name']}@{c['domain']}" for c in valid_cookies[:6]])
                except Exception:
                    sample = str(len(valid_cookies))
                logger.info(f"Injected {len(valid_cookies)} cookies for authentication [{sample}]")
                
                # Create page and navigate
                page = context.new_page()
                
                # Strategic navigation: establish authentication context
                try:
                    # First, visit the Microsoft login page to establish the authentication context
                    login_url = f"https://{login_domain}/common/oauth2/v2.0/authorize"
                    logger.info("Visiting Microsoft login domain to establish auth context")
                    page.goto(login_url, wait_until="domcontentloaded", timeout=20000)
                    page.wait_for_timeout(3000)
                    
                    # Check if we're already authenticated by looking for user profile or mail redirect
                    current_url = page.url
                    if "outlook.office.com" in current_url or "mail" in current_url:
                        logger.info("Already authenticated - redirected to Outlook")
                    else:
                        logger.info("Still on login page - cookies may need time to take effect")
                        page.wait_for_timeout(2000)
                        
                        # Try to trigger authentication by visiting a protected resource
                        try:
                            logger.info("Attempting to access protected resource to trigger authentication")
                            protected_url = f"{base_url}/mail/"
                            page.goto(protected_url, wait_until="domcontentloaded", timeout=15000)
                            page.wait_for_timeout(3000)
                        except Exception as e:
                            logger.warning(f"Protected resource access failed: {e}")
                        
                except Exception as e:
                    logger.warning(f"Login domain navigation failed: {e}")

                # Navigate to Outlook mail with proper authentication
                mail_url = f"{base_url}/mail/"
                logger.info(f"Navigating to Outlook mail: {mail_url}")
                page.goto(mail_url, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)
                
                # Check authentication status
                current_url = page.url.lower()
                logger.info(f"Current URL after navigation: {current_url}")
                
                # Check for authentication indicators - more specific for Outlook
                auth_indicators = [
                    "login.microsoftonline.com",
                    "login.live.com", 
                    "signin",
                    "oauth2",
                    "login"
                ]
                
                # Check if we're on a login page (but exclude valid Outlook URLs)
                is_on_login_page = any(indicator in current_url for indicator in auth_indicators)
                
                # Special case: if we're on outlook.office.com, check for specific login indicators
                if "outlook.office.com" in current_url:
                    # These are valid Outlook URLs that indicate successful authentication
                    valid_outlook_indicators = [
                        "msalauthredirect=true",  # Successful auth redirect
                        "mail",  # Mail interface
                        "compose",  # Compose interface
                        "inbox"  # Inbox interface
                    ]
                    
                    # If we have valid Outlook indicators, we're authenticated
                    has_valid_outlook_indicators = any(indicator in current_url for indicator in valid_outlook_indicators)
                    
                    if has_valid_outlook_indicators:
                        logger.info(f"Successfully authenticated - valid Outlook URL: {current_url}")
                        is_on_login_page = False  # Override the login page check
                    else:
                        # Check for actual login indicators in Outlook URL
                        login_indicators_in_outlook = [
                            "signin",
                            "login",
                            "auth",
                            "oauth2"
                        ]
                        is_on_login_page = any(indicator in current_url for indicator in login_indicators_in_outlook)
                
                if is_on_login_page:
                    logger.warning(f"Authentication check failed - redirected to login: {current_url}")
                    return False, "Cookies are invalid or expired - authentication required"
                
                # Additional check: look for Outlook-specific elements
                try:
                    # Wait for Outlook interface to load
                    page.wait_for_timeout(3000)
                    
                    # Check for Outlook-specific elements
                    outlook_elements = [
                        '[data-testid="compose-button"]',
                        '[aria-label*="compose"]',
                        '[aria-label*="New message"]',
                        'button[title*="compose"]',
                        'button[title*="New"]',
                        '.ms-Button--primary',
                        '[data-automation-id="compose-button"]',
                        '[data-testid="mail-module"]',  # Mail module
                        '[data-testid="inbox"]',  # Inbox
                        '.ms-CommandBar',  # Command bar
                        '[role="main"]'  # Main content area
                    ]
                    
                    found_outlook_element = False
                    for selector in outlook_elements:
                        try:
                            element = page.locator(selector).first
                            if element.is_visible(timeout=2000):
                                found_outlook_element = True
                                logger.info(f"Found Outlook interface element: {selector}")
                                break
                        except:
                            continue
                    
                    # If we're on a valid Outlook URL but no interface elements found,
                    # it might still be loading or the interface might be different
                    if not found_outlook_element and "outlook.office.com" in current_url:
                        logger.info("Outlook interface elements not found, but on valid Outlook URL - assuming authenticated")
                        found_outlook_element = True
                    
                    if not found_outlook_element:
                        logger.warning("Outlook interface not detected - may not be authenticated")
                        return False, "Outlook interface not accessible - authentication may have failed"
                        
                except Exception as e:
                    logger.warning(f"Error checking Outlook interface: {e}")
                    # Continue anyway - the URL check above is the primary indicator
                
                # Try to find compose interface
                    page.wait_for_timeout(3000)
                
                # Look for compose elements
                compose_selectors = [
                    '[data-testid="compose-button"]',
                    '[aria-label*="compose"]',
                    '[aria-label*="New message"]',
                    'button[title*="compose"]',
                    'button[title*="New"]',
                    '.ms-Button--primary',
                    '[role="button"]'
                ]
                
                compose_found = False
                for selector in compose_selectors:
                    try:
                        element = page.wait_for_selector(selector, timeout=5000)
                        if element:
                            compose_found = True
                            logger.info(f"Found compose element with selector: {selector}")
                            break
                    except:
                        continue
                
                if compose_found:
                    logger.info(f"Cookie authentication successful for {email} - compose interface found")
                    
                    # SAVE BROWSER SESSION for future use
                    try:
                        from pathlib import Path
                        sessions_dir = Path("sessions").resolve()
                        sessions_dir.mkdir(exist_ok=True)
                        safe_email = email.replace("@", "_at_").replace(".", "_")
                        session_file = sessions_dir / f"office365_{safe_email}.json"
                        
                        # Save the browser context state
                        context.storage_state(path=str(session_file))
                        logger.info(f"Saved browser session to: {session_file}")
                        
                        # Best-effort update of cookie manager metadata (may not exist yet)
                        try:
                            cookie_manager.update_account_status(email, 'active')
                            if getattr(cookie_manager, 'accounts', None) and email in cookie_manager.accounts:
                                cookie_manager.accounts[email]['session_file'] = str(session_file)
                                cookie_manager.accounts[email]['last_used'] = datetime.now().isoformat()
                        except Exception as meta_error:
                            logger.debug(f"Skipping cookie manager metadata update: {meta_error}")
                        
                    except Exception as session_error:
                        logger.warning(f"Could not save browser session: {session_error}")
                    
                    return True, "Authentication successful - compose interface accessible"
                else:
                    # Try to navigate to compose directly
                    compose_url = f"{base_url}/mail/deeplink/compose"
                    try:
                        page.goto(compose_url, wait_until="domcontentloaded", timeout=15000)

                        page.wait_for_timeout(3000)
                        compose_url_check = page.url.lower()
                        
                        if "compose" in compose_url_check or "mail" in compose_url_check:
                            logger.info(f"Cookie authentication successful for {email} - compose page accessible")
                            
                            # SAVE BROWSER SESSION for future use
                            try:
                                from pathlib import Path
                                sessions_dir = Path("sessions").resolve()
                                sessions_dir.mkdir(exist_ok=True)
                                safe_email = email.replace("@", "_at_").replace(".", "_")
                                session_file = sessions_dir / f"office365_{safe_email}.json"
                                
                                # Save the browser context state
                                context.storage_state(path=str(session_file))
                                logger.info(f"Saved browser session to: {session_file}")
                                
                                # Best-effort update of cookie manager metadata (may not exist yet)
                                try:
                                    cookie_manager.update_account_status(email, 'active')
                                    if getattr(cookie_manager, 'accounts', None) and email in cookie_manager.accounts:
                                        cookie_manager.accounts[email]['session_file'] = str(session_file)
                                        cookie_manager.accounts[email]['last_used'] = datetime.now().isoformat()
                                except Exception as meta_error:
                                    logger.debug(f"Skipping cookie manager metadata update: {meta_error}")
                                
                            except Exception as session_error:
                                logger.warning(f"Could not save browser session: {session_error}")
                            
                            return True, "Authentication successful - compose page accessible"
                        else:
                            logger.warning(f"Compose page check failed - URL: {compose_url_check}")
                            return False, "Cannot access compose page - cookies may be invalid"
                    except Exception as compose_error:
                        logger.warning(f"Error accessing compose page: {compose_error}")
                        return False, f"Error accessing compose page: {str(compose_error)}"
                
            except Exception as e:
                logger.error(f"Cookie authentication test failed for {email}: {e}")
                return False, f"Authentication test failed: {str(e)}"
            finally:
                try:
                    if 'page' in locals() and page:
                        page.close()
                except Exception:
                    pass
                try:
                    if 'browser' in locals():
                        browser.close()
                except Exception:
                    pass
                
    except Exception as e:
        logger.error(f"Error testing cookie authentication for {email}: {e}")
        return False, f"Authentication error: {str(e)}"


@app.route('/accounts/office365/paid', methods=['POST'])
@login_required
def add_paid_office365_account():
    """Add a paid Office 365 account with cookie data"""
    try:
        data = request.get_json()
        email = data.get('email')
        cookie_data = data.get('cookie_data')
        
        if not email or not cookie_data:
            return jsonify({'success': False, 'message': 'Email and cookie data are required'}), 400
        
        # FIRST: Test authentication with cookies
        logger.info(f"Testing authentication for {email} with cookies...")
        auth_success, auth_error = test_cookie_authentication(email, cookie_data)
        
        if not auth_success:
            logger.warning(f"Authentication failed for {email}: {auth_error}")
            return jsonify({'success': False, 'message': f'Authentication failed: {auth_error}'}), 400
        
        # Authentication successful - now save the account
        logger.info(f"Authentication successful for {email}, saving account...")
        
        # Add to cookie manager
        success = cookie_manager.add_account(email, cookie_data, 'paid')
        
        if success:
            # Also add to database for consistency (only if not already exists)
            try:
                # Check if account already exists
                session_db = db_manager.get_session()
                existing_account = session_db.query(EmailAccount).filter(EmailAccount.email == email).first()
                
                if not existing_account:
                    db_manager.add_email_account(email=email, provider='office365')
                    logger.info(f"Added paid Office 365 account to database: {email}")
                else:
                    logger.info(f"Office 365 account already exists in database: {email}")
                    
            except Exception as db_error:
                logger.warning(f"Added to cookie manager but failed to add to database: {db_error}")
            
            return jsonify({'success': True, 'message': 'Paid Office 365 account added successfully after authentication test'})
        else:
            return jsonify({'success': False, 'message': 'Failed to save account after successful authentication'}), 500
            
    except Exception as e:
        logger.error(f"Error adding paid Office 365 account: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500


if __name__ == '__main__':
    # Run the application
    app.run(
        host=config.HOST,
        port=config.PORT,
        debug=config.DEBUG
    )
