"""
Core email sending functionality with multithreading and rate limiting
Handles concurrent email sending across multiple providers and accounts
"""

import asyncio
import time
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime, timedelta
from dataclasses import dataclass
from enum import Enum
import aiohttp
from loguru import logger

from auth.base_auth import TokenManager, TokenExpiredError, OAuth2Error
from automation.office365_fast import fast_send_email, fast_send_email_with_cookies
from automation.office365_fast import _storage_path  # prefer saved session when available
from auth.office365_auth import Office365Auth
from auth.gmail_auth import GmailAuth
from auth.yahoo_auth import YahooAuth
from auth.hotmail_auth import HotmailAuth
from email_service.cookie_manager import cookie_manager
import database
from database import EmailAccount, EmailRecipient, EmailLog, EmailCampaign, get_db
from config import config
from ai.placeholder_replacer import placeholder_replacer

class EmailStatus(Enum):
    """Email sending status"""
    PENDING = "pending"
    SENDING = "sending"
    SENT = "sent"
    FAILED = "failed"
    THROTTLED = "throttled"
    RETRYING = "retrying"

@dataclass
class EmailTask:
    """Email sending task"""
    recipient: EmailRecipient
    account: EmailAccount
    subject: str
    body_html: str
    bcc_emails: List[str] = None
    attachments: List[Dict[str, Any]] = None
    custom_data: Dict[str, Any] = None
    proxy: str = None  # Add proxy field
    retry_count: int = 0
    max_retries: int = 5

@dataclass
class SendResult:
    """Email sending result"""
    success: bool
    status: EmailStatus
    error_message: Optional[str] = None
    response_data: Optional[Dict[str, Any]] = None
    processing_time: Optional[float] = None
    retry_after: Optional[int] = None

class RateLimiter:
    """Rate limiter for email sending"""
    
    def __init__(self, max_requests: int, time_window: int = 60):
        self.max_requests = max_requests
        self.time_window = time_window
        self.requests = []
        self._lock = asyncio.Lock()
    
    async def acquire(self) -> bool:
        """Acquire permission to send email"""
        async with self._lock:
            now = time.time()
            # Remove old requests outside time window
            self.requests = [req_time for req_time in self.requests 
                           if now - req_time < self.time_window]
            
            if len(self.requests) < self.max_requests:
                self.requests.append(now)
                return True
            return False
    
    async def wait_for_slot(self) -> float:
        """Wait for available slot and return wait time"""
        while not await self.acquire():
            await asyncio.sleep(1)
        return 0

class ProviderManager:
    """Manages email providers and their rate limits"""
    
    def __init__(self):
        self.providers = {
            'office365': Office365Auth(),
            'gmail': GmailAuth(),
            'yahoo': YahooAuth(),
            'hotmail': HotmailAuth()
        }
        
        # Rate limiters per provider
        self.rate_limiters = {
            provider: RateLimiter(
                max_requests=config.RATE_LIMIT_PER_MINUTE,
                time_window=60
            ) for provider in self.providers.keys()
        }
        
        # Semaphores for concurrency control
        self.semaphores = {
            provider: asyncio.Semaphore(config.MAX_CONCURRENT_PER_PROVIDER)
            for provider in self.providers.keys()
        }
    
    def get_provider(self, provider_name: str):
        """Get provider instance"""
        return self.providers.get(provider_name.lower())
    
    def get_rate_limiter(self, provider_name: str) -> RateLimiter:
        """Get rate limiter for provider"""
        return self.rate_limiters.get(provider_name.lower())
    
    def get_semaphore(self, provider_name: str) -> asyncio.Semaphore:
        """Get semaphore for provider"""
        return self.semaphores.get(provider_name.lower())

class EmailSender:
    """Main email sending class with multithreading support"""
    
    def __init__(self, db_manager):
        self.db_manager = db_manager
        self.token_manager = TokenManager(db_manager)
        self.provider_manager = ProviderManager()
        self.global_semaphore = asyncio.Semaphore(config.MAX_CONCURRENT_ACCOUNTS)
        # Cooperative cancellation per campaign
        self.cancel_events: Dict[int, asyncio.Event] = {}
        
        # Statistics
        self.stats = {
            'total_sent': 0,
            'total_failed': 0,
            'total_throttled': 0,
            'start_time': None,
            'active_tasks': 0
        }

    def request_cancel(self, campaign_id: int) -> None:
        """Signal cooperative cancellation for a running campaign."""
        evt = self.cancel_events.get(campaign_id)
        if evt is not None:
            evt.set()
    
    async def send_single_email(self, task: EmailTask) -> SendResult:
        """Send a single email"""
        start_time = time.time()
        
        try:
            access_token = None
            # Try to get a valid access token (OAuth path)
            try:
                access_token = await self.token_manager.get_valid_token(
                    task.account.id, task.account.provider
                )
            except Exception as token_error:
                # We'll consider automation fallback for Office365 below
                access_token = None
            
            # Get provider and rate limiter
            provider = self.provider_manager.get_provider(task.account.provider)
            rate_limiter = self.provider_manager.get_rate_limiter(task.account.provider)
            
            if not provider or not rate_limiter:
                return SendResult(
                    success=False,
                    status=EmailStatus.FAILED,
                    error_message=f"Provider {task.account.provider} not supported"
                )
            
            # Render placeholders per send (supports RAND/date/PROJECT_AI)
            try:
                # Use batch_rand from custom_data if provided (e.g., mass-send batches)
                batch_rand = None
                if task.custom_data and isinstance(task.custom_data, dict):
                    batch_rand = task.custom_data.get('batch_rand')
                rendered_subject = await placeholder_replacer.replace_placeholders(task.subject, batch_rand=batch_rand, tz_name="America/New_York")
                rendered_body = await placeholder_replacer.replace_placeholders(task.body_html, batch_rand=batch_rand, tz_name="America/New_York")
            except Exception:
                rendered_subject = task.subject
                rendered_body = task.body_html

            # Wait for rate limit slot
            await rate_limiter.wait_for_slot()
            
            # If we have an OAuth token (gmail/yahoo/hotmail, or office365 with OAuth), use provider API
            if access_token is not None and task.account.provider != 'office365':
                response = await provider.send_email(
                    access_token=access_token,
                    sender_email=task.account.email,
                    to_emails=[task.recipient.email],
                    subject=rendered_subject,
                    body_html=rendered_body,
                    bcc_emails=task.bcc_emails,
                    attachments=task.attachments
                )
            else:
                # Fallback path: Use Office365 fast browser automation when OAuth is not available
                if task.account.provider == 'office365':
                    # Prefer saved browser session (Playwright storage state) if present
                    try:
                        storage = _storage_path(task.account.email)
                        has_session = storage.exists()
                    except Exception:
                        has_session = False

                    if has_session:
                        send_ok = await asyncio.to_thread(
                            fast_send_email,
                            task.account.email,
                            [task.account.email],
                            rendered_subject,
                            rendered_body,
                            task.bcc_emails or [],
                            task.attachments or [],
                            True,
                            task.proxy
                        )
                        if not send_ok:
                            # Fallback to cookies if session-based send fails
                            cookies = cookie_manager.get_cookies_for_injection(task.account.email)
                            if cookies and cookie_manager.is_cookie_valid(task.account.email):
                                send_ok = await asyncio.to_thread(
                                    fast_send_email_with_cookies,
                                    task.account.email,
                                    [task.account.email],
                                    rendered_subject,
                                    rendered_body,
                                    cookies,
                                    task.bcc_emails or [],
                                    task.attachments or [],
                                    True,
                                    task.proxy
                                )
                                if not send_ok:
                                    raise Exception("Office365 send failed (session+cookies)")
                                response = { 'method': 'cookie_injection', 'provider': 'office365', 'ok': True }
                                cookie_manager.update_account_status(task.account.email, 'active')
                            else:
                                raise Exception("Office365 automation send failed (no valid cookies)")
                        else:
                            response = { 'method': 'automation_session', 'provider': 'office365', 'ok': True }
                    else:
                        # No session: try cookies first, then automation without saved session
                        cookies = cookie_manager.get_cookies_for_injection(task.account.email)
                        if cookies and cookie_manager.is_cookie_valid(task.account.email):
                            send_ok = await asyncio.to_thread(
                                fast_send_email_with_cookies,
                                task.account.email,
                                [task.account.email],
                                rendered_subject,
                                rendered_body,
                                cookies,
                                task.bcc_emails or [],
                                task.attachments or [],
                                True,
                                task.proxy
                            )
                            if not send_ok:
                                raise Exception("Office365 cookie injection send failed")
                            response = { 'method': 'cookie_injection', 'provider': 'office365', 'ok': True }
                            cookie_manager.update_account_status(task.account.email, 'active')
                        else:
                            send_ok = await asyncio.to_thread(
                                fast_send_email,
                                task.account.email,
                                [task.account.email],
                                rendered_subject,
                                rendered_body,
                                task.bcc_emails or [],
                                task.attachments or [],
                                True,
                                task.proxy
                            )
                            if not send_ok:
                                raise Exception("Office365 automation send failed")
                            response = { 'method': 'automation', 'provider': 'office365', 'ok': True }
                else:
                    # Non-office365 without token cannot send
                    raise Exception("No valid access token for provider")
            
            processing_time = time.time() - start_time
            
            # Update account statistics
            self.db_manager.increment_account_stats(task.account.id, success=True)
            
            # Log successful send
            self.db_manager.log_email_send(
                account_email=task.account.email,
                recipient_email=task.recipient.email,
                subject=task.subject,
                status='success',
                response_data=response,
                processing_time=processing_time,
                campaign_id=task.recipient.campaign_id,
                recipient_id=task.recipient.id
            )
            
            return SendResult(
                success=True,
                status=EmailStatus.SENT,
                response_data=response,
                processing_time=processing_time
            )
            
        except TokenExpiredError:
            # Token expired, mark account as needing re-auth
            self.db_manager.get_session().query(EmailAccount).filter(
                EmailAccount.id == task.account.id
            ).update({'is_active': False})
            
            return SendResult(
                success=False,
                status=EmailStatus.FAILED,
                error_message="Token expired - account needs re-authentication"
            )
            
        except OAuth2Error as e:
            error_msg = str(e)
            status = EmailStatus.THROTTLED if "Rate limited" in error_msg else EmailStatus.FAILED
            retry_after = None
            
            if "Rate limited" in error_msg:
                # Extract retry after time
                try:
                    retry_after = int(error_msg.split("Retry after ")[1].split(" ")[0])
                except:
                    retry_after = 60
            
            # Update account statistics
            self.db_manager.increment_account_stats(task.account.id, success=False)
            
            # Log failed send
            self.db_manager.log_email_send(
                account_email=task.account.email,
                recipient_email=task.recipient.email,
                subject=task.subject,
                status='failed',
                error_message=error_msg,
                processing_time=time.time() - start_time,
                campaign_id=task.recipient.campaign_id,
                recipient_id=task.recipient.id
            )
            
            return SendResult(
                success=False,
                status=status,
                error_message=error_msg,
                retry_after=retry_after
            )
            
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            logger.error(f"Unexpected error sending email: {error_msg}")
            
            # Update account statistics
            self.db_manager.increment_account_stats(task.account.id, success=False)
            
            # Log failed send
            self.db_manager.log_email_send(
                account_email=task.account.email,
                recipient_email=task.recipient.email,
                subject=task.subject,
                status='failed',
                error_message=error_msg,
                processing_time=time.time() - start_time,
                campaign_id=task.recipient.campaign_id,
                recipient_id=task.recipient.id
            )
            
            return SendResult(
                success=False,
                status=EmailStatus.FAILED,
                error_message=error_msg
            )
    
    async def send_with_retry(self, task: EmailTask, cancel_event: asyncio.Event = None) -> SendResult:
        """Send email with retry logic"""
        for attempt in range(task.max_retries):
            task.retry_count = attempt
            # If cancellation requested, stop scheduling more attempts
            if cancel_event is not None and cancel_event.is_set():
                return SendResult(
                    success=False,
                    status=EmailStatus.FAILED,
                    error_message="Cancelled"
                )
            
            # Get semaphore for provider
            semaphore = self.provider_manager.get_semaphore(task.account.provider)
            
            async with semaphore:
                result = await self.send_single_email(task)
                
                if result.success:
                    return result
                
                # Handle retry logic
                if result.status == EmailStatus.THROTTLED and attempt < task.max_retries - 1:
                    wait_time = result.retry_after or (2 ** attempt)
                    logger.info(f"Throttled, waiting {wait_time} seconds before retry {attempt + 1}")
                    await asyncio.sleep(wait_time)
                    continue
                
                if result.status == EmailStatus.FAILED and attempt < task.max_retries - 1:
                    wait_time = config.RETRY_DELAY * (2 ** attempt)
                    logger.info(f"Failed, waiting {wait_time} seconds before retry {attempt + 1}")
                    await asyncio.sleep(wait_time)
                    continue
                
                # Max retries reached or non-retryable error
                return result
        
        return result
    
    async def send_campaign(self, campaign_id: int, max_concurrent: int = None, selected_accounts: list = None, proxy: str = None, batch_size: int = 50, email_list: list = None) -> Dict[str, Any]:
        """Send email campaign with multithreading, selected accounts, proxy, and batch processing"""
        if max_concurrent:
            self.global_semaphore = asyncio.Semaphore(max_concurrent)
        
        self.stats['start_time'] = datetime.utcnow()
        # Prepare cancellation event
        cancel_event = self.cancel_events.get(campaign_id)
        if cancel_event is None:
            cancel_event = asyncio.Event()
            self.cancel_events[campaign_id] = cancel_event
        
        # Get campaign and recipients
        campaign = self.db_manager.get_session().query(EmailCampaign).filter(
            EmailCampaign.id == campaign_id
        ).first()
        
        if not campaign:
            return {'success': False, 'error': 'Campaign not found'}
        
        # Use uploaded email list if provided, otherwise use database recipients
        if email_list:
            # Create virtual recipients from uploaded email list
            recipients = []
            for email in email_list:
                # Create a virtual recipient object
                class VirtualRecipient:
                    def __init__(self, email):
                        self.email = email
                        self.status = 'pending'
                        self.custom_data = {}
                
                recipients.append(VirtualRecipient(email))
            logger.info(f"Using uploaded email list with {len(recipients)} recipients")
        else:
            # Use database recipients (existing behavior)
            recipients = self.db_manager.get_session().query(EmailRecipient).filter(
                EmailRecipient.campaign_id == campaign_id,
                EmailRecipient.status == 'pending'
            ).all()
            logger.info(f"Using database recipients with {len(recipients)} recipients")
        
        if not recipients:
            return {'success': False, 'error': 'No recipients found'}
        
        # Get active accounts - use selected accounts if provided
        if selected_accounts:
            # Filter active accounts to only include selected ones
            all_active_accounts = self.db_manager.get_active_accounts()
            active_accounts = [acc for acc in all_active_accounts if acc.email in selected_accounts]
            logger.info(f"Using selected accounts: {[acc.email for acc in active_accounts]}")
        else:
            # Use all active accounts (default behavior)
            active_accounts = self.db_manager.get_active_accounts()
            logger.info(f"Using all active accounts: {[acc.email for acc in active_accounts]}")
        
        if not active_accounts:
            return {'success': False, 'error': 'No active accounts found'}
        
        # Update campaign status
        campaign.status = 'running'
        campaign.started_at = datetime.utcnow()
        campaign.total_recipients = len(recipients)
        self.db_manager.get_session().commit()
        
        # Create tasks with batch processing
        tasks = []
        account_index = 0
        
        if email_list and batch_size > 1:
            # Batch processing for uploaded email lists
            for i in range(0, len(recipients), batch_size):
                batch = recipients[i:i + batch_size]
                account = active_accounts[account_index % len(active_accounts)]
                account_index += 1
                
                # Create BCC list for this batch
                bcc_emails = [recipient.email for recipient in batch]
                
                # Create a single task for the entire batch
                task = EmailTask(
                    recipient=batch[0],  # Use first recipient as primary
                    account=account,
                    subject=campaign.subject,
                    body_html=campaign.body_html,
                    custom_data={},
                    proxy=proxy,
                    bcc_emails=bcc_emails  # Add BCC support
                )
                tasks.append(task)
                logger.info(f"Created batch task with {len(bcc_emails)} BCC recipients")
        else:
            # Individual email processing (existing behavior)
            for recipient in recipients:
                account = active_accounts[account_index % len(active_accounts)]
                account_index += 1
                
                task = EmailTask(
                    recipient=recipient,
                    account=account,
                    subject=campaign.subject,
                    body_html=campaign.body_html,
                    custom_data=getattr(recipient, 'custom_data', {}),
                    proxy=proxy
                )
                tasks.append(task)
        
        # Send emails concurrently
        logger.info(f"Starting campaign {campaign_id} with {len(tasks)} emails")
        
        async def send_task(task: EmailTask):
            async with self.global_semaphore:
                self.stats['active_tasks'] += 1
                try:
                    # Early exit if cancelled (won't interrupt in-flight sends)
                    if cancel_event.is_set():
                        return
                    result = await self.send_with_retry(task, cancel_event)
                    
                    # Update recipient status
                    if result.success:
                        task.recipient.status = 'sent'
                        task.recipient.sent_at = datetime.utcnow()
                        task.recipient.account_used = task.account.email
                        self.stats['total_sent'] += 1
                    else:
                        task.recipient.status = 'failed'
                        task.recipient.error_message = result.error_message
                        self.stats['total_failed'] += 1
                        
                        if result.status == EmailStatus.THROTTLED:
                            self.stats['total_throttled'] += 1
                    
                    # Update campaign statistics
                    if result.success:
                        campaign.sent_count += 1
                    else:
                        campaign.failed_count += 1
                    
                    self.db_manager.get_session().commit()
                    
                finally:
                    self.stats['active_tasks'] -= 1
        
        # Execute all tasks
        await asyncio.gather(*[send_task(task) for task in tasks], return_exceptions=True)
        
        # Update campaign status
        if cancel_event.is_set():
            campaign.status = 'cancelled'
        else:
            campaign.status = 'completed'
            campaign.completed_at = datetime.utcnow()
        self.db_manager.get_session().commit()
        
        # Calculate final statistics
        total_time = (datetime.utcnow() - self.stats['start_time']).total_seconds()
        
        return {
            'success': True,
            'campaign_id': campaign_id,
            'total_recipients': len(recipients),
            'sent': self.stats['total_sent'],
            'failed': self.stats['total_failed'],
            'throttled': self.stats['total_throttled'],
            'total_time': total_time,
            'emails_per_minute': (self.stats['total_sent'] / total_time * 60) if total_time > 0 else 0
        }
    
    async def send_bulk_emails(self, emails: List[Dict[str, Any]], 
                              max_concurrent: int = None) -> Dict[str, Any]:
        """Send bulk emails without campaign structure"""
        if max_concurrent:
            self.global_semaphore = asyncio.Semaphore(max_concurrent)
        
        self.stats['start_time'] = datetime.utcnow()
        
        # Get active accounts
        active_accounts = self.db_manager.get_active_accounts()
        if not active_accounts:
            return {'success': False, 'error': 'No active accounts found'}
        
        # Create tasks
        tasks = []
        account_index = 0
        
        for email_data in emails:
            # Round-robin account selection
            account = active_accounts[account_index % len(active_accounts)]
            account_index += 1
            
            # Create mock recipient for logging
            recipient = EmailRecipient(
                id=0,  # Temporary ID
                campaign_id=0,
                email=email_data['to'],
                name=email_data.get('name'),
                custom_data=email_data.get('custom_data')
            )
            
            task = EmailTask(
                recipient=recipient,
                account=account,
                subject=email_data['subject'],
                body_html=email_data['body_html'],
                bcc_emails=email_data.get('bcc'),
                attachments=email_data.get('attachments'),
                custom_data=email_data.get('custom_data')
            )
            tasks.append(task)
        
        # Send emails concurrently
        logger.info(f"Starting bulk send with {len(tasks)} emails")
        
        results = []
        async def send_task(task: EmailTask):
            async with self.global_semaphore:
                self.stats['active_tasks'] += 1
                try:
                    result = await self.send_with_retry(task)
                    results.append({
                        'email': task.recipient.email,
                        'success': result.success,
                        'error': result.error_message,
                        'processing_time': result.processing_time
                    })
                    
                    if result.success:
                        self.stats['total_sent'] += 1
                    else:
                        self.stats['total_failed'] += 1
                        
                        if result.status == EmailStatus.THROTTLED:
                            self.stats['total_throttled'] += 1
                    
                finally:
                    self.stats['active_tasks'] -= 1
        
        # Execute all tasks
        await asyncio.gather(*[send_task(task) for task in tasks], return_exceptions=True)
        
        # Calculate final statistics
        total_time = (datetime.utcnow() - self.stats['start_time']).total_seconds()
        
        return {
            'success': True,
            'total_emails': len(emails),
            'sent': self.stats['total_sent'],
            'failed': self.stats['total_failed'],
            'throttled': self.stats['total_throttled'],
            'total_time': total_time,
            'emails_per_minute': (self.stats['total_sent'] / total_time * 60) if total_time > 0 else 0,
            'results': results
        }
    
    def get_stats(self) -> Dict[str, Any]:
        """Get current sending statistics"""
        return {
            **self.stats,
            'active_tasks': self.stats['active_tasks'],
            'uptime': (datetime.utcnow() - self.stats['start_time']).total_seconds() 
                     if self.stats['start_time'] else 0
        }
    
    def reset_stats(self):
        """Reset statistics"""
        self.stats = {
            'total_sent': 0,
            'total_failed': 0,
            'total_throttled': 0,
            'start_time': None,
            'active_tasks': 0
        }
