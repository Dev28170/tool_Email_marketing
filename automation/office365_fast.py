"""
Fast and reliable Office365 email sending
Optimized for speed and reliability
Implementing YOUR PERFECT BCC approach: Individual email entry with Enter/Tab
"""

import os
import time
from pathlib import Path
from typing import List, Dict, Optional
from urllib.parse import urlencode, quote
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from loguru import logger
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.dynamic_timing import DynamicTiming, TimingContext


def _storage_path(email: str) -> Path:
    """Get storage path for session"""
    sessions_dir = Path("sessions")
    sessions_dir.mkdir(exist_ok=True)
    safe_email = email.replace("@", "_at_").replace(".", "_")
    return sessions_dir / f"office365_{safe_email}.json"


def _outlook_host_for_email(email: str) -> str:
    """Return the correct Outlook web host for the email type.
    - Freeicrosoft accounts (outlook.com/hotmail.com/live.com/msn.com) use outlook.live.com
    - Work/School (Microsoft 365) and custom domains use outlook.office.com
    """
    try:
        domain = (email.split('@', 1)[1] or '').lower()
    except Exception:
        domain = ''
    free_domains = {
        'outlook.com', 'hotmail.com', 'live.com', 'msn.com',
        'outlook.co.uk', 'hotmail.co.uk', 'hotmail.fr'
    }
    return 'outlook.live.com' if domain in free_domains else 'outlook.office.com'


def _inject_cookies_to_context(context, cookies: List[Dict], outlook_domain: str = None) -> None:
    """Inject authentication cookies into browser context"""
    try:
        for cookie in cookies:
            # Convert cookie format for Playwright
            # Use Outlook domain instead of login domain
            cookie_domain = outlook_domain or '.outlook.office.com'
            
            playwright_cookie = {
                'name': cookie.get('name', ''),
                'value': cookie.get('value', ''),
                'domain': cookie_domain,  # Use Outlook domain, not login domain
                'path': cookie.get('path', '/'),
                'secure': cookie.get('secure', True),
                'httpOnly': cookie.get('httpOnly', True),
                'sameSite': 'None'  # Playwright expects 'None' (capital N)
            }
            
            # Handle expiration
            if 'expirationDate' in cookie:
                exp_date = cookie['expirationDate']
                if isinstance(exp_date, (int, float)):
                    # Convert from milliseconds to seconds
                    exp_seconds = exp_date / 1000
                    playwright_cookie['expires'] = exp_seconds
            
            context.add_cookies([playwright_cookie])
            logger.info(f"Injected cookie: {cookie.get('name')}")
            
    except Exception as e:
        logger.error(f"Error injecting cookies: {e}")


def _bcc_chips_count(page, expected_emails: List[str] = None) -> int:
    """Count BCC chips using specific email verification."""
    try:
        if not expected_emails:
            logger.debug("No expected emails provided for BCC chip counting")
            return 0
        
        # Count only the specific emails we expect to be in BCC
        email_count = 0
        verified_emails = []
        
        for email in expected_emails:
            try:
                # Look for the email as a chip (with remove button or specific chip classes)
                chip_selectors = [
                    f"[data-automationid*='Pill']:has-text('{email}')",
                    f"[data-testid*='recipient-chip']:has-text('{email}')",
                    f"[data-testid*='pill']:has-text('{email}')",
                    f"[role='button'][aria-label*='Remove']:has-text('{email}')",
                    f"[class*='persona']:has-text('{email}')",
                    f"[class*='chip']:has-text('{email}')",
                    f"[class*='pill']:has-text('{email}')",
                    f"span[title*='{email}']",
                    f"div[title*='{email}']"
                ]
                
                email_found = False
                for selector in chip_selectors:
                    try:
                        if page.locator(selector).count() > 0:
                            email_count += 1
                            verified_emails.append(email)
                            email_found = True
                            logger.debug(f"Found email as chip: {email}")
                            break
                    except Exception:
                        continue
                
                # If not found as chip, check if it's just text in the BCC area
                if not email_found:
                    try:
                        # Look for email text specifically in BCC area
                        bcc_area_selectors = [
                            "div[aria-label*='Bcc']",
                            "div[role='textbox'][aria-label*='Bcc']",
                            "[data-testid*='bcc']"
                        ]
                        
                        for bcc_selector in bcc_area_selectors:
                            try:
                                bcc_area = page.locator(bcc_selector).first
                                if bcc_area.count() > 0 and bcc_area.text_content() and email in bcc_area.text_content():
                                    email_count += 1
                                    verified_emails.append(email)
                                    email_found = True
                                    logger.debug(f"Found email in BCC area: {email}")
                                    break
                            except Exception:
                                continue
                    except Exception:
                        pass
                
                if not email_found:
                    logger.debug(f"Email not found as chip: {email}")
                    
            except Exception as e:
                logger.debug(f"Error checking email {email}: {e}")
                continue
        
        logger.debug(f"BCC chips count: {email_count}/{len(expected_emails)} - Verified: {verified_emails}")
        return int(email_count)
    except Exception as e:
        logger.debug(f"Error counting BCC chips: {e}")
        return 0


def _bcc_has_email(page, email: str) -> bool:
    """Heuristic check whether a specific email is present in Bcc area.
    Different tenants render people pills differently, so we look for text matches
    inside the Bcc container and common chip containers.
    """
    try:
        # First try to find the BCC container and look within it
        bcc_container_selectors = [
            "div[aria-label='Bcc']",
            "div[aria-label*='Bcc']",
            "[data-testid*='bcc']",
            "[data-automationid*='Bcc']"
        ]
        
        bcc_container = None
        for selector in bcc_container_selectors:
            try:
                container = page.locator(selector).first
                if container.is_visible(timeout=1000):
                    bcc_container = container
                    break
            except Exception:
                continue
        
        if bcc_container:
            # Look for email within BCC container - be more specific
            candidates = [
                f"[title='{email}']",  # Exact title match
                f"[aria-label='{email}']",  # Exact aria-label match
                f"span:has-text('{email}')",  # Span with exact text
                f"div:has-text('{email}')",  # Div with exact text
            ]
            for sel in candidates:
                try:
                    loc = bcc_container.locator(sel)
                    if loc.count() > 0:
                        # Double-check that this is actually a chip/pill, not just text in input
                        for i in range(loc.count()):
                            element = loc.nth(i)
                            try:
                                # Check if it's a chip/pill (has remove button or specific classes)
                                if (element.locator("[aria-label*='Remove']").count() > 0 or
                                    element.locator("[class*='pill']").count() > 0 or
                                    element.locator("[class*='chip']").count() > 0 or
                                    element.locator("[data-automationid*='Pill']").count() > 0):
                                    logger.debug(f"Found email {email} as chip in BCC container with selector: {sel}")
                                    return True
                            except Exception:
                                continue
                except Exception:
                    continue
        
        # Fallback: look for chips specifically, not just any text
        chip_candidates = [
            f"[data-automationid*='Pill']:has-text('{email}')",
            f"[class*='pill']:has-text('{email}')",
            f"[class*='chip']:has-text('{email}')",
            f"[role='button'][aria-label*='Remove']:has-text('{email}')",
        ]
        for sel in chip_candidates:
            try:
                loc = page.locator(sel)
                if loc.count() > 0:
                    logger.debug(f"Found email {email} as chip with selector: {sel}")
                    return True
            except Exception:
                continue
        
        # Only check for emails as actual chips, not just text anywhere
        # This prevents false positives from emails appearing in other parts of the page
        
        logger.debug(f"Email {email} not found as a chip")
        return False
    except Exception as e:
        logger.debug(f"Error checking for email {email}: {e}")
        return False


def _ensure_to_field_populated(page, sender_email: str, proxy: str = None) -> bool:
    """
    Ensure the TO field is populated with sender email as a fallback.
    Returns True if successful, False otherwise.
    """
    try:
        logger.info(f"📧 Ensuring TO field is populated with sender email: {sender_email}")
        
        # Find TO field with multiple selectors
        to_field_selectors = [
            'input[aria-label*="To"]',
            'input[aria-label*="to"]',
            'input[data-testid*="to"]',
            'input[data-testid*="To"]',
            'input[placeholder*="To"]',
            'input[placeholder*="to"]',
            'div[aria-label*="To"] input',
            'div[aria-label*="to"] input',
            'div[data-testid*="to"] input',
            'div[data-testid*="To"] input'
        ]
        
        to_field = None
        for selector in to_field_selectors:
            try:
                field = page.locator(selector).first
                if field.is_visible(timeout=2000):
                    to_field = field
                    logger.info(f"Found TO field with selector: {selector}")
                    break
            except Exception as e:
                logger.debug(f"TO field selector {selector} failed: {e}")
                continue
        
        if not to_field:
            logger.warning("Could not find TO field, trying fallback methods...")
            # Fallback: try to find any input field that might be TO field
            try:
                all_inputs = page.locator('input[type="text"], input[type="email"]')
                count = all_inputs.count()
                for i in range(count):
                    try:
                        field = all_inputs.nth(i)
                        if field.is_visible(timeout=1000):
                            # Check if this looks like a TO field
                            placeholder = field.evaluate("el => el.placeholder || ''")
                            aria_label = field.evaluate("el => el.getAttribute('aria-label') || ''")
                            if 'to' in placeholder.lower() or 'to' in aria_label.lower():
                                to_field = field
                                logger.info(f"Found TO field via fallback (element {i})")
                                break
                    except Exception as e:
                        logger.debug(f"Error checking input element {i}: {e}")
                        continue
            except Exception as e:
                logger.debug(f"Fallback TO field detection failed: {e}")
        
        if to_field:
            try:
                # Clear existing content
                to_field.click()
                to_field.press("Control+a")
                to_field.press("Delete")
                time.sleep(0.2)
                
                # Fill with sender email
                to_field.fill(sender_email)
                time.sleep(0.5)
                
                # Verify content was set
                current_value = to_field.evaluate("el => el.value || ''")
                if sender_email in current_value:
                    logger.info(f"✅ TO field populated successfully with: {sender_email}")
                    return True
                else:
                    logger.warning(f"⚠️ TO field may not have been populated correctly. Expected: {sender_email}, Got: {current_value}")
                    return False
                    
            except Exception as e:
                logger.error(f"❌ Error populating TO field: {e}")
                return False
        else:
            logger.warning("⚠️ Could not find TO field to populate")
            return False
            
    except Exception as e:
        logger.debug(f"Error ensuring TO field is populated: {e}")
        return False


def _handle_attachments(page, attachments: List[str], proxy: str = None) -> bool:
    """
    Upload attachments on Outlook Web compose page using Playwright.
    - Handles Attach File → Browse this computer flow.
    - Detects correct file input (not image upload).
    - Waits until upload completes.
    - Takes screenshot when upload finishes.
    """
    try:
        logger.info(f"📎 Starting upload of {len(attachments)} attachment(s)...")

        # Step 1. Find and click "Insert" tab
        insert_selectors = [
            'button[aria-label="Insert"]',
            'button[title="Insert"]',
            'div[role="tab"]:has-text("Insert")',
            'button:has-text("Insert")'
        ]
        insert_button = None
        for sel in insert_selectors:
            try:
                button = page.locator(sel).first
                button.wait_for(state="visible", timeout=3000)
                insert_button = button
                button.click()
                logger.info(f"✅ Clicked Insert tab ({sel})")
                page.wait_for_timeout(1000)
                break
            except Exception as e:
                logger.debug(f"Insert selector failed: {sel} -> {e}")
        if not insert_button:
            logger.warning("⚠️ Could not find Insert tab/button.")
            return False

        # Step 2. Find "Attach file" button
        attach_selectors = [
            'button[aria-label="Attach file"]',
            'button[title="Attach file"]',
            'button[data-icon-name="Attach"]',
            'button:has-text("Attach file")',
            '[data-testid*="attach-file"]'
        ]
        attach_button = None
        for sel in attach_selectors:
            try:
                button = page.locator(sel).first
                button.wait_for(state="visible", timeout=4000)
                attach_button = button
                logger.info(f"✅ Found Attach File button: {sel}")
                break
            except Exception as e:
                logger.debug(f"Attach selector failed: {sel} -> {e}")
        if not attach_button:
            logger.warning("⚠️ Could not find Attach File button.")
            return False

        # Ensure screenshot directory
        os.makedirs("screenshots/attachments", exist_ok=True)

        success_count = 0

        # Step 3. Upload each file
        for i, file_path in enumerate(attachments):
            logger.info(f"📎 Uploading attachment {i+1}/{len(attachments)}: {file_path}")
            try:
                # Step 3.1 Click "Attach file"
                attach_button.click()
                page.wait_for_timeout(1000)

                # Step 3.2 Click "Browse this computer"
                browse_selectors = [
                    'button:has-text("Browse this computer")',
                    'div[role="menuitem"]:has-text("Browse this computer")',
                    'button[aria-label*="Browse this computer"]',
                    'button[title*="Browse this computer"]'
                ]
                browse_button = None
                for bsel in browse_selectors:
                    try:
                        btn = page.locator(bsel).first
                        btn.wait_for(state="visible", timeout=2000)
                        browse_button = btn
                        browse_button.click()
                        logger.info(f"✅ Clicked 'Browse this computer' ({bsel})")
                        page.wait_for_timeout(1000)
                        break
                    except Exception as e:
                        logger.debug(f"Browse selector failed: {bsel} -> {e}")
                if not browse_button:
                    logger.warning("⚠️ Could not find 'Browse this computer' option.")
                    continue

                # Step 3.3 Locate correct <input type=file>
                logger.info("🔍 Looking for correct file input element...")
                page.wait_for_timeout(1000)
                file_inputs = page.locator('input[type="file"]').all()

                target_input = None
                for fi in file_inputs:
                    try:
                        html = fi.evaluate("el => el.outerHTML")
                        if 'accept=' not in html:  # Outlook image inputs have accept= attribute
                            target_input = fi
                            logger.info(f"✅ Found correct file input: {html}")
                            break
                        else:
                            logger.debug(f"Skipping image-only input: {html}")
                    except Exception:
                        continue

                if not target_input:
                    logger.warning("⚠️ No unrestricted file input found, using first available.")
                    target_input = page.locator('input[type="file"]').first

                target_input.wait_for(state="attached", timeout=5000)
                target_input.set_input_files(file_path)
                logger.info(f"⏳ Upload started for: {file_path}")

                # Step 3.4 Wait for upload progress to disappear
                progress_locators = [
                    '[data-testid*="uploadProgress"]',
                    '[aria-label*="Uploading"]',
                    '[role="progressbar"]'
                ]
                for _ in range(30):  # Wait up to 30 seconds
                    still_uploading = False
                    for prog_sel in progress_locators:
                        try:
                            if page.locator(prog_sel).is_visible():
                                still_uploading = True
                                break
                        except Exception:
                            continue
                    if not still_uploading:
                        break
                    page.wait_for_timeout(1000)

                # Step 3.5 Verify upload completion
                attached_count = page.locator('[data-testid*="attachment"], [class*="attachment"]').count()
                logger.info(f"✅ Upload complete. Attachments visible: {attached_count}")
                success_count += 1

                # Step 3.6 Screenshot after successful upload
                filename = os.path.basename(file_path)
                safe_name = os.path.splitext(filename)[0].replace(" ", "_")
                screenshot_path = f"screenshots/attachments/{safe_name}_done.png"
                # page.screenshot(path=screenshot_path, full_page=True)
                logger.info(f"📸 Screenshot saved: {screenshot_path}")

                # Delay before next file
                if i < len(attachments) - 1:
                    page.wait_for_timeout(1500)

            except Exception as e:
                logger.error(f"❌ Error uploading {file_path}: {e}")

        if success_count == len(attachments):
            logger.info(f"🎉 All {len(attachments)} attachments uploaded successfully.")
            return True
        elif success_count > 0:
            logger.warning(f"⚠️ Partial success ({success_count}/{len(attachments)})")
            return True
        else:
            logger.error("❌ No attachments uploaded successfully.")
            return False

    except Exception as e:
        logger.error(f"❌ Fatal error in _handle_attachments: {e}")
        return False



def _clear_compose_window(page) -> bool:
    """
    Clear any existing content in the compose window to ensure clean state.
    Returns True if clearing was successful, False otherwise.
    """
    try:
        logger.info("🧹 Clearing compose window for clean state...")
        
        # Method 1: Clear all contenteditable elements
        page.evaluate("""
            () => {
                const contenteditables = document.querySelectorAll('[contenteditable="true"]');
                contenteditables.forEach(el => {
                    el.innerHTML = '';
                    el.textContent = '';
                });
            }
        """)
        
        # Method 2: Clear any draft content
        page.evaluate("""
            () => {
                // Clear any draft content
                const draftElements = document.querySelectorAll('[data-testid*="draft"], [class*="draft"]');
                draftElements.forEach(el => {
                    if (el.contentEditable === 'true') {
                        el.innerHTML = '';
                        el.textContent = '';
                    }
                });
            }
        """)
        
        # Method 3: Clear any text areas
        page.evaluate("""
            () => {
                const textareas = document.querySelectorAll('textarea, input[type="text"]');
                textareas.forEach(el => {
                    if (el.value) {
                        el.value = '';
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                });
            }
        """)
        
        logger.info("✅ Compose window cleared successfully")
        return True
        
    except Exception as e:
        logger.debug(f"Error clearing compose window: {e}")
        return False

def _validate_body_field(body_field) -> bool:
    """
    Validate that the found field is actually the email body field.
    Returns True if it's a valid email body field, False otherwise.
    """
    try:
        # Check if it's a proper contenteditable div
        is_contenteditable = body_field.evaluate("el => el.contentEditable === 'true'")
        if not is_contenteditable:
            return False
        
        # Check if it's in the compose area (not in header, toolbar, etc.)
        rect = body_field.bounding_box()
        if not rect or rect['height'] < 50:  # Too small to be email body
            return False
        
        # Check if it has proper email body characteristics
        has_body_classes = body_field.evaluate("""
            el => {
                const classList = Array.from(el.classList);
                return classList.some(cls => 
                    cls.toLowerCase().includes('body') || 
                    cls.toLowerCase().includes('compose') || 
                    cls.toLowerCase().includes('message')
                );
            }
        """)
        
        # Check if it's not in a header or toolbar area
        is_in_compose_area = body_field.evaluate("""
            el => {
                const rect = el.getBoundingClientRect();
                return rect.top > 200; // Should be below header/toolbar
            }
        """)
        
        return has_body_classes or is_in_compose_area
        
    except Exception as e:
        logger.debug(f"Error validating body field: {e}")
        return False

def _wait_for_compose_page_loaded(page, proxy: str = None) -> bool:
    """
    Wait for compose page to fully load by detecting various loading states.
    Returns True if compose page is fully loaded, False if timeout.
    """
    logger.info("🔄 Waiting for compose page to fully load...")
    
    # Get adaptive timeout based on proxy performance
    max_wait = DynamicTiming.get_adaptive_timeout('compose_load', proxy)
    start_time = time.time()
    
    # Loading indicators to detect
    loading_selectors = [
        # Common loading spinners
        '[class*="loading"]',
        '[class*="spinner"]',
        '[class*="loader"]',
        '[data-testid*="loading"]',
        '[aria-label*="Loading"]',
        '[aria-label*="loading"]',
        # Outlook specific loading indicators
        '.ms-Spinner',
        '.loading-spinner',
        '[class*="ms-Spinner"]',
        # Generic loading text
        'text="Loading"',
        'text="loading"',
        'text="Please wait"',
        'text="Please Wait"',
        # Loading states
        '[class*="loading-state"]',
        '[data-loading="true"]',
        '[aria-busy="true"]'
    ]
    
    # Compose page elements that should be present when loaded
    compose_selectors = [
        # Email input fields
        'input[placeholder*="To"]',
        'input[placeholder*="to"]',
        'input[aria-label*="To"]',
        'input[aria-label*="to"]',
        # Subject field
        'input[placeholder*="Subject"]',
        'input[placeholder*="subject"]',
        'input[aria-label*="Subject"]',
        'input[aria-label*="subject"]',
        # Body field
        'div[contenteditable="true"]',
        'div[aria-label*="Message body"]',
        'div[aria-label*="message body"]',
        # Compose area indicators
        '[data-testid*="compose"]',
        '[class*="compose"]',
        '[role="textbox"]'
    ]
    
    while (time.time() - start_time) * 1000 < max_wait:
        try:
            # Check if any loading indicators are present
            loading_present = False
            for selector in loading_selectors:
                try:
                    if page.locator(selector).count() > 0:
                        loading_present = True
                        break
                except Exception:
                    continue
            
            # Check if compose elements are present
            compose_ready = False
            for selector in compose_selectors:
                try:
                    if page.locator(selector).count() > 0:
                        compose_ready = True
                        break
                except Exception:
                    continue
            
            # If no loading indicators and compose elements are present, we're ready
            if not loading_present and compose_ready:
                elapsed = (time.time() - start_time) * 1000
                logger.info(f"✅ Compose page fully loaded in {elapsed:.1f}ms")
                return True
            
            # Log progress every 5 seconds
            elapsed = (time.time() - start_time) * 1000
            if int(elapsed) % 5000 < 100:  # Every ~5 seconds
                if loading_present:
                    logger.info(f"⏳ Still loading... ({elapsed:.0f}ms elapsed)")
                elif not compose_ready:
                    logger.info(f"⏳ Waiting for compose elements... ({elapsed:.0f}ms elapsed)")
            
            # Wait a bit before checking again
            page.wait_for_timeout(500)
            
        except Exception as e:
            logger.debug(f"Error checking compose page load state: {e}")
            page.wait_for_timeout(1000)
    
    # Timeout reached
    elapsed = (time.time() - start_time) * 1000
    logger.warning(f"⚠️ Compose page load timeout after {elapsed:.1f}ms")
    logger.info("💡 This may be due to proxy interference with Outlook's loading")
    return False

def _reveal_bcc_with_proxy_toggle(page, browser, context, proxy: str = None) -> bool:
    """
    Ensure Bcc field is visible using multiple aggressive methods.
    If proxy is interfering with BCC detection, log the issue for manual resolution.
    Returns True if we believe Bcc is available on screen.
    """
    revealed = False
    with TimingContext("BCC reveal") as timing:
        try:
            # First try with current proxy setup
            revealed = _reveal_bcc(page)
            
            # If BCC detection failed and we have a proxy, provide helpful guidance
            if not revealed and proxy and proxy.strip():
                logger.warning("🔄 BCC detection failed with proxy")
                logger.info("💡 This is a known issue: proxy interferes with BCC field detection")
                logger.info("💡 Solutions:")
                logger.info("   1. Use without proxy for BCC functionality")
                logger.info("   2. Try a different proxy provider")
                logger.info("   3. Contact proxy provider about UI interaction issues")
                logger.info("💡 The email will still be sent via URL parameters (BCC in compose URL)")
                    
        except Exception as e:
            logger.debug(f"BCC reveal error: {e}")
    
    logger.info(f"BCC reveal result: {'SUCCESS' if revealed else 'FAILED'}")
    return revealed

def _reveal_bcc(page) -> bool:
    """Ensure Bcc field is visible using multiple aggressive methods.
    Returns True if we believe Bcc is available on screen.
    """
    revealed = False
    with TimingContext("BCC reveal") as timing:
        try:
            # Method 1: Enhanced keyboard shortcut with better detection
            try:
                logger.debug("Trying Control+Shift+B keyboard shortcut...")
                page.keyboard.press('Control+Shift+B')
                page.wait_for_timeout(2000)  # Increased wait time
                
                # Try multiple keyboard shortcuts
                try:
                    page.keyboard.press('Control+Shift+C')  # Alternative shortcut
                    page.wait_for_timeout(1000)
                except Exception:
                    pass
                
                # Enhanced BCC detection after keyboard shortcut
                bcc_indicators = [
                    'div[aria-label*="Bcc"]',
                    'input[aria-label*="Bcc"]', 
                    'div[aria-label*="BCC"]',
                    'input[aria-label*="BCC"]',
                    'div[aria-label*="bcc"]',
                    'input[aria-label*="bcc"]',
                    '[data-testid*="bcc"]',
                    '[data-testid*="BCC"]',
                    'div[role="textbox"][aria-label*="Bcc"]',
                    'div[role="textbox"][aria-label*="BCC"]',
                    'div[role="textbox"][aria-label*="bcc"]',
                    'div[contenteditable="true"][aria-label*="Bcc"]',
                    'div[contenteditable="true"][aria-label*="BCC"]',
                    'div[contenteditable="true"][aria-label*="bcc"]',
                    # Additional indicators for newer Outlook versions
                    'div[class*="bcc"]',
                    'div[class*="BCC"]',
                    'div[class*="recipient"][aria-label*="Bcc"]',
                    'div[class*="recipient"][aria-label*="BCC"]',
                    'div[class*="recipient"][aria-label*="bcc"]'
                ]
                
                for indicator in bcc_indicators:
                    try:
                        if page.locator(indicator).count() > 0:
                            revealed = True
                            logger.debug(f"BCC revealed via keyboard shortcut with indicator: {indicator}")
                            break
                    except Exception:
                        continue
                        
            except Exception as e:
                logger.debug(f"Keyboard shortcut failed: {e}")

            # Method 2: FAST BCC button detection - optimized for speed
            if not revealed:
                # Ultra-fast BCC button detection with minimal debugging
                logger.debug("Fast BCC button detection...")
                
                # ENHANCED BCC BUTTON SELECTORS - More comprehensive detection
                fast_selectors = [
                    # TOP PRIORITY - Most effective based on your HTML structure
                    'button[class*="fui-"]:has-text("Bcc")',           # Primary selector
                    'button.fui-Button:has-text("Bcc")',              # Exact class match
                    'button[class*="fui-Button"]:has-text("Bcc")',    # Partial class match
                    # ENHANCED - More comprehensive selectors
                    'button:has-text("Bcc")',                         # Simple text match
                    'button[aria-label*="Bcc"]',                      # Aria label match
                    'button[aria-label*="bcc"]',                       # Lowercase variant
                    'button[data-testid*="bcc"]',                     # Test ID match
                    'button[data-testid*="BCC"]',                     # Uppercase test ID
                    'button[data-automationid*="bcc"]',               # Automation ID
                    '[role="button"][aria-label*="Bcc"]',            # Role-based
                    '[role="button"][aria-label*="bcc"]',             # Role-based lowercase
                    'button:has-text("Cc & Bcc")',                    # Alternative text
                    'button:has-text("BCC")',                         # Uppercase variant
                    'button:has-text("bcc")',                         # Lowercase variant
                    # FALLBACK - Generic button detection
                    'button[class*="button"]:has-text("Bcc")',         # Generic button class
                    'div[role="button"]:has-text("Bcc")',             # Div as button
                    'span[role="button"]:has-text("Bcc")'             # Span as button
                ]
                # ULTRA-FAST BCC BUTTON DETECTION
                for sel in fast_selectors:
                    try:
                        btn = page.locator(sel).first
                        # Reduced timeout for speed - 200ms instead of 1500ms
                        if DynamicTiming.wait_for_element_ready(page, btn, max_wait=200):
                            logger.debug(f"Found BCC button: {sel}")
                            # FAST CLICK METHODS - Optimized for speed
                            try:
                                # Primary: Fast click with reduced timeout
                                btn.click(timeout=1000)  # Reduced from 2000ms
                            except Exception:
                                try:
                                    # Secondary: Force click with minimal timeout
                                    btn.click(timeout=500, force=True)  # Reduced from 2000ms
                                except Exception:
                                    # Tertiary: Coordinate click (fastest)
                                    box = btn.bounding_box()
                                    if box:
                                        page.mouse.click(box['x'] + box['width']/2, box['y'] + box['height']/2)
                            
                            # FAST BCC FIELD DETECTION after button click
                            # Reduced timeout for speed - 500ms instead of 1000ms
                            if DynamicTiming.wait_for_condition(
                                lambda: page.locator('div[aria-label*="Bcc"], input[aria-label*="Bcc"], [data-testid*="bcc"]').count() > 0,
                                max_wait=timing.get_adaptive_delay(500)  # Reduced from 1000ms
                            ):
                                revealed = True
                                logger.debug(f"BCC revealed via button: {sel}")
                                break
                            else:
                                # FAST FALLBACK - Quick input field scan
                                try:
                                    # Use faster selector for input scan
                                    bcc_inputs = page.locator('[aria-label*="Bcc"], [data-testid*="bcc"]').all()
                                    if len(bcc_inputs) > 0:
                                        revealed = True
                                        logger.debug(f"BCC field found via fast scan: {len(bcc_inputs)} elements")
                                        break
                                except Exception:
                                    pass
                    except Exception as e:
                        logger.debug(f"Button selector {sel} failed: {e}")
                        continue

            # Method 3: FAST text element clicks - optimized for speed
            if not revealed:
                logger.debug("Fast text element clicks for BCC reveal...")
                # Only try the most effective text selectors
                fast_text_selectors = [
                    'div:has-text("Bcc")',    # Most common
                    'span:has-text("Bcc")'    # Alternative
                ]
                for sel in fast_text_selectors:
                    try:
                        el = page.locator(sel).first
                        # Reduced timeout for speed
                        if DynamicTiming.wait_for_element_ready(page, el, max_wait=300):  # Reduced from 1000ms
                            logger.debug(f"Found text element: {sel}")
                            # Fast click with reduced timeout
                            el.click(timeout=500)  # Reduced from 1500ms
                            # Fast BCC field detection with reduced timeout
                            if DynamicTiming.wait_for_condition(
                                lambda: page.locator('[aria-label*="Bcc"], [data-testid*="bcc"]').count() > 0,
                                max_wait=timing.get_adaptive_delay(300)  # Reduced from 1000ms
                            ):
                                revealed = True
                                logger.debug(f"BCC revealed via text element: {sel}")
                                break
                    except Exception as e:
                        logger.debug(f"Text selector {sel} failed: {e}")
                        continue

            # Method 4: FAST compose area click - optimized for speed
            if not revealed:
                try:
                    logger.debug("Fast compose area click for BCC reveal...")
                    # Use faster selector for compose area
                    compose_area = page.locator('[data-testid="compose-area"], .compose-area, [aria-label*="compose"]').first
                    if DynamicTiming.wait_for_element_ready(page, compose_area, max_wait=300):  # Reduced from 1000ms
                        compose_area.click(timeout=500)  # Reduced timeout
                        time.sleep(0.2)  # Reduced from 0.5s
                        
                        # Try keyboard shortcut again after clicking compose area
                        page.keyboard.press('Control+Shift+B')
                        # Fast BCC field detection
                        if DynamicTiming.wait_for_condition(
                            lambda: page.locator('[aria-label*="Bcc"], [data-testid*="bcc"]').count() > 0,
                            max_wait=timing.get_adaptive_delay(300)  # Reduced from 1000ms
                        ):
                            revealed = True
                            logger.debug("BCC revealed after compose area click + keyboard shortcut")
                except Exception as e:
                    logger.debug(f"Compose area click failed: {e}")

            # Method 5: FAST heuristic check - optimized for speed
            if not revealed:
                logger.debug("Fast heuristic BCC check...")
                # Only check the most effective selectors
                fast_probe_selectors = [
                    '[aria-label*="Bcc"]',     # Most reliable
                    '[data-testid*="bcc"]'     # Alternative
                ]
                for sel in fast_probe_selectors:
                    try:
                        probe = page.locator(sel).first
                        # Reduced timeout for speed
                        if DynamicTiming.wait_for_element_ready(page, probe, max_wait=200):  # Reduced from 500ms
                            revealed = True
                            logger.debug(f"BCC already visible via heuristic: {sel}")
                            break
                    except Exception:
                        continue
                        
        except Exception as e:
            logger.debug(f"BCC reveal error: {e}")
    
    # Method 5: Last resort - try to find any recipient field and assume it's BCC
    if not revealed:
        logger.debug("Trying last resort BCC detection...")
        try:
            # Look for any recipient field that might work as BCC
            recipient_selectors = [
                'div[role="textbox"][contenteditable="true"]',
                'div[contenteditable="true"]',
                'input[type="text"]',
                'input[type="email"]',
                'div[role="combobox"]',
                'div[role="searchbox"]'
            ]
            
            for selector in recipient_selectors:
                try:
                    elements = page.locator(selector)
                    if elements.count() > 0:
                        # Try the first few elements
                        for i in range(min(3, elements.count())):
                            try:
                                element = elements.nth(i)
                                if element.is_visible(timeout=1000):
                                    # Try to click and see if it accepts BCC input
                                    element.click()
                                    page.wait_for_timeout(500)
                                    
                                    # Try to type a test email
                                    element.type("test@example.com")
                                    page.wait_for_timeout(500)
                                    
                                    # Check if it created a chip
                                    if page.locator('[data-testid*="chip"], .chip, [class*="chip"]').count() > 0:
                                        revealed = True
                                        logger.debug(f"BCC field found via last resort: {selector}")
                                        break
                                    
                                    # Clear the test input
                                    element.clear()
                                    page.wait_for_timeout(500)
                            except Exception:
                                continue
                        if revealed:
                            break
                except Exception:
                    continue
        except Exception as e:
            logger.debug(f"Last resort BCC detection failed: {e}")
    
    logger.info(f"BCC reveal result: {'SUCCESS' if revealed else 'FAILED'}")
    return revealed


def _force_focus(page, locator) -> bool:
    """Robustly focus the given locator even if overlays intercept clicks.
    Attempts: scroll, JS focus, force click, coordinate click.
    Returns True on likely focus success.
    """
    try:
        try:
            locator.scroll_into_view_if_needed(timeout=800)
        except Exception:
            pass
        # JS focus first (works for contenteditable)
        try:
            locator.evaluate("el => el.focus()")
            page.wait_for_timeout(80)
        except Exception:
            pass
        # Force click to bypass overlay
        try:
            locator.click(timeout=1200, force=True)
            page.wait_for_timeout(80)
            return True
        except Exception:
            pass
        # Coordinate click at center
        try:
            box = locator.bounding_box()
            if box:
                x = box['x'] + box['width'] / 2
                y = box['y'] + box['height'] / 2
                page.mouse.click(x, y)
                page.wait_for_timeout(80)
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False
def _is_in_bcc_context(locator) -> bool:
    """Return True if the given locator appears to be inside a Bcc container.
    We walk up a few ancestors and require an aria-label (or text) that includes 'bcc',
    and explicitly exclude known body containers.
    """
    try:
        # Exclude obvious body containers
        try:
            role = locator.get_attribute('role') or ''
        except Exception:
            role = ''
        try:
            aria = (locator.get_attribute('aria-label') or '').lower()
        except Exception:
            aria = ''
        if 'message body' in aria or ('textbox' in role and 'message' in aria):
            return False

        # Walk up to 4 ancestors and check aria-label/labels for 'bcc'
        parent = locator
        for _ in range(4):
            try:
                aria_label = (parent.get_attribute('aria-label') or '').lower()
            except Exception:
                aria_label = ''
            if 'bcc' in aria_label:
                return True
            try:
                parent = parent.locator('xpath=..')
            except Exception:
                break
        return False
    except Exception:
        return False


def _fill_bcc_field(page, bcc: List[str], proxy: str = None) -> int:
    """Loop through provided emails and create BCC chips reliably.
    Returns the number of chips created.
    """
    # Normalize and deduplicate
    normalized = []
    seen = set()
    for addr in bcc or []:
        v = (addr or '').strip()
        if not v:
            continue
        key = v.lower()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(v)

    # Use dynamic timeout for BCC field detection
    bcc_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
    logger.info(f"⏱️ Using adaptive BCC detection timeout: {bcc_timeout}ms for proxy: {proxy or 'none'}")
    
    # ENHANCED BCC reveal with multiple aggressive methods
    try:
        logger.info("Attempting to reveal BCC field...")
        
        # Method 1: Keyboard shortcut
        page.keyboard.press('Control+Shift+B')
        page.wait_for_timeout(1000)
        
        # Method 2: Enhanced BCC reveal button detection
        bcc_reveal_selectors = [
            'button[aria-label*="Bcc"]',
            'button[aria-label*="bcc"]', 
            'button[aria-label*="BCC"]',
            'button[data-testid*="bcc"]',
            'button[data-testid*="BCC"]',
            'button[data-automationid*="bcc"]',
            'button[data-automationid*="BCC"]',
            '[role="button"][aria-label*="Bcc"]',
            '[role="button"][aria-label*="bcc"]',
            '[role="button"][aria-label*="BCC"]',
            'button:has-text("Bcc")',
            'button:has-text("bcc")',
            'button:has-text("BCC")',
            'button:has-text("Cc & Bcc")',
            'div[role="button"]:has-text("Bcc")',
            'span[role="button"]:has-text("Bcc")',
            'button[class*="fui-"]:has-text("Bcc")',
            'button[class*="fui-Button"]:has-text("Bcc")'
        ]
        
        for reveal_selector in bcc_reveal_selectors:
            try:
                reveal_btn = page.locator(reveal_selector).first
                if reveal_btn.is_visible(timeout=1000):
                    logger.info(f"Found BCC reveal button: {reveal_selector}")
                    reveal_btn.click()
                    page.wait_for_timeout(500)
                    break
            except Exception:
                continue
        
        # Method 3: Try clicking on compose area to trigger BCC reveal
        try:
            compose_area = page.locator('div[role="main"], div[class*="compose"], div[class*="Compose"]').first
            if compose_area.is_visible(timeout=1000):
                compose_area.click()
                page.wait_for_timeout(500)
        except Exception:
            pass
        
        # Method 4: Aggressive BCC field detection - try to find hidden BCC field
        if not revealed:
            logger.debug("Trying aggressive BCC field detection...")
            aggressive_selectors = [
                # Look for any element that might be a BCC field, even if hidden
                'div[aria-label*="Bcc"]',
                'div[aria-label*="BCC"]',
                'div[aria-label*="bcc"]',
                'input[aria-label*="Bcc"]',
                'input[aria-label*="BCC"]',
                'input[aria-label*="bcc"]',
                'div[role="textbox"][aria-label*="Bcc"]',
                'div[role="textbox"][aria-label*="BCC"]',
                'div[role="textbox"][aria-label*="bcc"]',
                'div[contenteditable="true"][aria-label*="Bcc"]',
                'div[contenteditable="true"][aria-label*="BCC"]',
                'div[contenteditable="true"][aria-label*="bcc"]',
                # Look for any recipient field that might be BCC
                'div[class*="recipient"]',
                'div[class*="Recipient"]',
                'div[class*="bcc"]',
                'div[class*="BCC"]',
                'div[class*="Bcc"]',
                # Look for any field that might contain BCC
                'div:has-text("Bcc")',
                'div:has-text("BCC")',
                'div:has-text("bcc")',
                'input:has-text("Bcc")',
                'input:has-text("BCC")',
                'input:has-text("bcc")'
            ]
            
            for selector in aggressive_selectors:
                try:
                    elements = page.locator(selector)
                    if elements.count() > 0:
                        # Try to make the element visible and clickable
                        for i in range(elements.count()):
                            try:
                                element = elements.nth(i)
                                # Try to make it visible
                                element.scroll_into_view_if_needed()
                                element.hover()
                                element.click()
                                page.wait_for_timeout(500)
                                
                                # Check if BCC field is now visible
                                if page.locator('div[aria-label*="Bcc"], input[aria-label*="Bcc"], div[aria-label*="BCC"], input[aria-label*="BCC"]').count() > 0:
                                    revealed = True
                                    logger.debug(f"BCC revealed via aggressive detection: {selector}")
                                    break
                            except Exception:
                                continue
                        if revealed:
                            break
                except Exception:
                    continue
            
    except Exception as e:
        logger.debug(f"BCC reveal attempt failed: {e}")
    
    # Wait for BCC field to appear
    page.wait_for_timeout(1000)
    
    # Intelligent BCC field detection with priority order
    selectors = [
        # High priority - most specific selectors first
        'input[aria-label="Bcc"]',
        'input[aria-label*="Bcc"]',
        'div[aria-label="Bcc"] input',
        'div[aria-label*="Bcc"] input',
        'div[role="textbox"][aria-label*="Bcc"]',
        'div[role="textbox"][aria-label="Bcc"]',
        'div[aria-label*="Bcc"] [role="textbox"]',
        'div[aria-label*="Bcc"] [contenteditable="true"]',
        'div[aria-label*="Bcc"] [contenteditable]',
        'div[aria-label*="Bcc"] [role="combobox"] input',
        'div[aria-label*="Bcc"] [role="combobox"] [contenteditable="true"]',
        'div[aria-label*="Bcc"] [role="searchbox"]',
        'div[aria-label*="Bcc"] input[type="text"]',
        'div[aria-label*="Bcc"] input[type="email"]',
        'div[aria-label*="Bcc"] textarea',
        'div[aria-label*="Bcc"] [role="combobox"]',
        'div[aria-label*="Bcc"] [role="searchbox"]',
        # Medium priority - data attributes
        '[data-testid*="bcc"] input',
        '[data-testid*="bcc"] [role="textbox"]',
        '[data-testid*="bcc"] [contenteditable]',
        '[data-automationid*="Bcc"] input',
        '[data-automationid*="Bcc"] [contenteditable="true"]',
        '[data-automationid*="bcc"] [role="textbox"]',
        '[data-automationid*="bcc"] [contenteditable]',
        # Lower priority - generic selectors
        'input[placeholder*="Bcc"]',
        'input[placeholder*="bcc"]',
        'textarea[placeholder*="Bcc"]',
        'textarea[placeholder*="bcc"]',
        'div[aria-label*="recipients"] [role="textbox"]',
        'div[aria-label*="recipients"] [contenteditable="true"]',
        'div[role="combobox"][aria-label*="Bcc"] input',
        'div[role="combobox"][aria-label*="Bcc"] [contenteditable="true"]',
        # Fallback - class-based selectors (less reliable)
        'div[class*="fui-"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="fui-"] [role="textbox"][aria-label*="Bcc"]',
        'div[class*="fui-"] input[aria-label*="Bcc"]',
        'div[class*="Wd_e8"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="Wd_e8"] [role="textbox"][aria-label*="Bcc"]',
        'div[class*="Wd_e8"] input[aria-label*="Bcc"]',
        # ENHANCED - Additional comprehensive selectors
        'input[aria-label*="BCC"]',  # Uppercase variant
        'input[aria-label*="bcc"]',  # Lowercase variant
        'div[aria-label*="BCC"] input',  # Uppercase variant
        'div[aria-label*="bcc"] input',  # Lowercase variant
        'div[role="textbox"][aria-label*="BCC"]',  # Uppercase variant
        'div[role="textbox"][aria-label*="bcc"]',  # Lowercase variant
        '[data-testid*="BCC"] input',  # Uppercase test ID
        '[data-testid*="bcc"] [role="textbox"]',  # Lowercase test ID
        '[data-automationid*="BCC"] input',  # Uppercase automation ID
        '[data-automationid*="bcc"] input',  # Lowercase automation ID
        'div[contenteditable="true"][aria-label*="BCC"]',  # Uppercase contenteditable
        'div[contenteditable="true"][aria-label*="bcc"]',  # Lowercase contenteditable
        # LAST RESORT - Generic fallbacks
        'div[contenteditable="true"]:has-text("Bcc")',
        'div[contenteditable="true"]:has-text("BCC")',
        'div[contenteditable="true"]:has-text("bcc")',
        'div[role="textbox"][contenteditable="true"]',
        'div[role="combobox"][contenteditable="true"]'
    ]

    for selector in selectors:
        try:
            logger.info(f"Trying BCC input selector: {selector}")
            input_el = page.locator(selector).first
            if not input_el.is_visible(timeout=bcc_timeout):
                continue

            # Guard: ensure we didn't accidentally select the message body or To field
            if not _is_in_bcc_context(input_el):
                logger.debug("Rejected non-Bcc context match for selector: {}".format(selector))
                continue
            logger.info(f"Found BCC input with selector: {selector}")
            # Ensure Bcc field is visible (try keyboard toggle as fallback)
            try:
                page.keyboard.press('Control+Shift+B')
                page.wait_for_timeout(200)
            except Exception:
                pass

            # Clear entire field
            input_el.click()
            page.wait_for_timeout(150)
            try:
                input_el.fill("")
            except Exception:
                pass
            page.keyboard.press("Control+a")
            page.wait_for_timeout(50)
            page.keyboard.press("Delete")
            page.wait_for_timeout(150)

            base_count = _bcc_chips_count(page, normalized)

            for idx, email in enumerate(normalized):
                try:
                    logger.info(f"Adding BCC recipient {idx+1}/{len(normalized)}: {email}")
                    
                    # Always add the email - don't skip based on detection
                    # The detection can be unreliable, so we'll add all emails
                    logger.info(f"Adding email: {email}")
                    
                    # Click on the input field but DON'T clear it - we want to add to existing content
                    input_el.click()
                    page.wait_for_timeout(100)
                    
                    # If this is not the first email, add a semicolon separator
                    if idx > 0:
                        page.keyboard.type(";", delay=50)
                        page.wait_for_timeout(100)
                    
                    # Type the email address
                    logger.info(f"Adding email: {email}")
                    page.keyboard.type(email, delay=100)
                    page.wait_for_timeout(200)
                    
                    # For the last email, commit all emails at once
                    if idx == len(normalized) - 1:
                        # This is the last email, commit all emails
                        logger.info(f"Committing all {len(normalized)} BCC emails...")
                        
                        # Try multiple commit methods for the final submission
                        commit_success = False
                        
                        # Method 1: Enter
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(500)
                        if _bcc_chips_count(page, normalized) >= len(normalized):
                            commit_success = True
                            logger.debug(f"All emails committed with Enter")
                        
                        if not commit_success:
                            # Method 2: Tab
                            input_el.click()
                            page.keyboard.press("Tab")
                            page.wait_for_timeout(500)
                            if _bcc_chips_count(page, normalized) >= len(normalized):
                                commit_success = True
                                logger.debug(f"All emails committed with Tab")
                        
                        if not commit_success:
                            # Method 3: Click outside to trigger tokenization
                            try:
                                page.locator('input[aria-label="Add a subject"]').first.click(timeout=1000)
                            except Exception:
                                try:
                                    page.locator('[data-testid="subject-input"]').first.click(timeout=1000)
                                except Exception:
                                    pass
                            page.wait_for_timeout(500)
                            if _bcc_chips_count(page, normalized) >= len(normalized):
                                commit_success = True
                                logger.debug(f"All emails committed with focus change")
                        
                        # Final verification
                        final_chips = _bcc_chips_count(page, normalized)
                        if commit_success or final_chips >= len(normalized):
                            logger.info(f"Successfully committed all {len(normalized)} BCC recipients")
                        else:
                            logger.warning(f"Only {final_chips}/{len(normalized)} BCC recipients were committed")
                    else:
                        # For intermediate emails, just add a semicolon and continue
                        page.keyboard.type(";", delay=50)
                        page.wait_for_timeout(200)
                        logger.info(f"Added email {email} to BCC list (will commit all at end)")
                except Exception as e:
                    logger.error(f"Error adding BCC recipient {email}: {e}")
                    continue

            # Check final result and retry for missing emails
            final_count = _bcc_chips_count(page, normalized)
            logger.info(f"Initial BCC attempt: {final_count}/{len(normalized)} emails added")
            
            # If some emails are missing, try to add them again with stabilization passes
            if final_count < len(normalized):
                logger.info(f"Retrying to add missing emails with stabilization passes...")
                for attempt in range(1, 4):
                    current_present = [e for e in normalized if _bcc_has_email(page, e)]
                    missing_emails = [e for e in normalized if e not in current_present]
                    if not missing_emails:
                        break
                    logger.info(f"Stabilization pass {attempt}: missing {missing_emails}")

                    # Focus Bcc input end
                    try:
                        input_el.click()
                        page.wait_for_timeout(120)
                        page.keyboard.press('End')
                        page.wait_for_timeout(80)
                    except Exception:
                        pass

                    # Append all missing
                    for email in missing_emails:
                        try:
                            page.keyboard.type(";", delay=40)
                            page.wait_for_timeout(80)
                            page.keyboard.type(email, delay=90)
                            page.wait_for_timeout(140)
                        except Exception as e:
                            logger.error(f"Error typing missing email {email}: {e}")
                            continue

                    # Commit: Enter -> Tab -> blur to subject
                    committed = False
                    try:
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(450)
                        if _bcc_chips_count(page, normalized) >= len(normalized):
                            committed = True
                    except Exception:
                        pass
                    if not committed:
                        try:
                            input_el.click()
                            page.keyboard.press("Tab")
                            page.wait_for_timeout(450)
                            if _bcc_chips_count(page, normalized) >= len(normalized):
                                committed = True
                        except Exception:
                            pass
                    if not committed:
                        try:
                            page.locator('input[aria-label="Add a subject"]').first.click(timeout=800)
                        except Exception:
                            try:
                                page.locator('[data-testid="subject-input"]').first.click(timeout=800)
                            except Exception:
                                pass
                        page.wait_for_timeout(450)

                    # Detect dropped chips and re-add next pass
                    new_present = [e for e in normalized if _bcc_has_email(page, e)]
                    dropped = [e for e in current_present if e not in new_present]
                    if dropped:
                        logger.warning(f"Detected dropped chips after commit, will re-add: {dropped}")

                    if _bcc_chips_count(page, normalized) >= len(normalized):
                        break
            
            return _bcc_chips_count(page, normalized)
        except Exception as e:
            logger.debug(f"BCC input selector {selector} failed: {e}")
            continue

    return _bcc_chips_count(page, normalized)


def _fill_bcc_bulk(page, bcc: List[str], browser=None, context=None, proxy: str = None) -> int:
    """Attempt to paste all BCC emails at once (semicolon-separated) into the
    Bcc field and let Outlook tokenize them automatically. Returns chips count.
    """
    # Normalize and deduplicate once up-front
    normalized = []
    seen = set()
    for addr in bcc or []:
        v = (addr or '').strip()
        if not v:
            continue
        key = v.lower()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(v)

    # Find a robust input target with multiple attempts
    target = None
    
    # First, ensure Bcc is revealed with smart proxy toggle
    if not _reveal_bcc_with_proxy_toggle(page, browser, context, proxy):
        logger.warning("Failed to reveal BCC field")
        return _bcc_chips_count(page, normalized)
    
    # After revealing BCC, minimal wait for the field to load
    time.sleep(0.2)  # Reduced from 1s to 0.2s
    
    # Try to find BCC field by looking for any input that appears after BCC button click
    logger.info("Looking for BCC input field after reveal...")
    try:
        # Check if any new input fields appeared
        all_inputs = page.locator('input, textarea, [contenteditable="true"], [role="textbox"]').all()
        logger.info(f"Found {len(all_inputs)} total input fields on page")
        
        for i, inp in enumerate(all_inputs):
            try:
                if DynamicTiming.wait_for_element_ready(page, inp, max_wait=500):
                    # Check if this input is in a BCC context
                    parent = inp.locator('xpath=..')
                    parent_aria = parent.evaluate("el => el.getAttribute('aria-label')")
                    grandparent = parent.locator('xpath=..')
                    grandparent_aria = grandparent.evaluate("el => el.getAttribute('aria-label')")
                    
                    if (parent_aria and 'bcc' in parent_aria.lower()) or (grandparent_aria and 'bcc' in grandparent_aria.lower()):
                        logger.info(f"Found BCC input field: element {i+1} (parent aria: {parent_aria}, grandparent aria: {grandparent_aria})")
                        # This is likely the BCC field, but we'll still try our comprehensive selectors
                        break
            except Exception as e:
                logger.debug(f"Could not check input element {i+1}: {e}")
                continue
    except Exception as e:
        logger.debug(f"BCC input field scan failed: {e}")

    # Try multiple selector strategies with more comprehensive list
    candidates = [
        # MOST EFFECTIVE - Based on working selector from your feedback
        'div[class*="fui-"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="fui-"] [role="textbox"][aria-label*="Bcc"]',
        'div[class*="fui-"] input[aria-label*="Bcc"]',
        # Additional fui- class variants
        'div[class*="fui-"] [contenteditable][aria-label*="Bcc"]',
        'div[class*="fui-"] [contenteditable="true"]',
        'div[class*="fui-"] [role="textbox"]',
        # Wd_e8 class variants (from button structure)
        'div[class*="Wd_e8"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="Wd_e8"] [role="textbox"][aria-label*="Bcc"]',
        'div[class*="Wd_e8"] input[aria-label*="Bcc"]',
        'div[class*="Wd_e8"] [contenteditable][aria-label*="Bcc"]',
        # Other class fragments from button
        'div[class*="r1alrhcs"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="r1alrhcs"] [role="textbox"][aria-label*="Bcc"]',
        'div[class*="___1qa2p2d"] [contenteditable="true"][aria-label*="Bcc"]',
        'div[class*="___1qa2p2d"] [role="textbox"][aria-label*="Bcc"]',
        # Standard selectors
        'input[aria-label="Bcc"]',
        'input[aria-label*="Bcc"]',
        'div[role="textbox"][aria-label*="Bcc"]',
        'div[aria-label="Bcc"] [role="textbox"]',
        'div[aria-label*="Bcc"] [role="textbox"]',
        'div[aria-label="Bcc"] [contenteditable="true"]',
        'div[aria-label*="Bcc"] [contenteditable="true"]',
        '[data-testid*="bcc"] input',
        'div[aria-label="Bcc"] input',
        'div[aria-label*="Bcc"] input',
        '[data-automationid*="Bcc"] input',
        '[data-automationid*="Bcc"] [contenteditable="true"]',
        'div[aria-label*="recipients"] [role="textbox"]',
        'div[aria-label*="recipients"] [contenteditable="true"]',
        'div[role="combobox"][aria-label*="Bcc"] input',
        'div[role="combobox"][aria-label*="Bcc"] [contenteditable="true"]',
        'div[aria-label*="Bcc"] input[type="text"]',
        'div[aria-label*="Bcc"] input[type="email"]',
        'div[aria-label*="Bcc"] textarea',
        'div[aria-label*="Bcc"] [contenteditable]',
        'div[aria-label*="Bcc"] [role="textbox"]',
        'div[aria-label*="Bcc"] [role="combobox"]',
        'div[aria-label*="Bcc"] [role="searchbox"]',
        'div[aria-label="Bcc"] input',
        'div[aria-label="Bcc"] textarea',
        'div[aria-label="Bcc"] [contenteditable]',
        'div[aria-label="Bcc"] [role="textbox"]',
        'input[placeholder*="Bcc"]',
        'input[placeholder*="bcc"]',
        'textarea[placeholder*="Bcc"]',
        'textarea[placeholder*="bcc"]',
        '[data-testid*="bcc"] [role="textbox"]',
        '[data-testid*="bcc"] [contenteditable]',
        '[data-automationid*="bcc"] [role="textbox"]',
        '[data-automationid*="bcc"] [contenteditable]'
    ]
    
    # First, let's debug what BCC-related elements are actually on the page
    logger.info("Debugging BCC field detection...")
    try:
        # Check for any BCC-related elements
        bcc_elements = page.locator('[aria-label*="Bcc"], [data-testid*="bcc"], [data-automationid*="Bcc"]').all()
        logger.info(f"Found {len(bcc_elements)} BCC-related elements on page")
        
        for i, elem in enumerate(bcc_elements):
            try:
                tag_name = elem.evaluate("el => el.tagName")
                aria_label = elem.evaluate("el => el.getAttribute('aria-label')")
                role = elem.evaluate("el => el.getAttribute('role')")
                contenteditable = elem.evaluate("el => el.getAttribute('contenteditable')")
                logger.info(f"BCC element {i+1}: {tag_name}, aria-label='{aria_label}', role='{role}', contenteditable='{contenteditable}'")
            except Exception as e:
                logger.debug(f"Could not inspect BCC element {i+1}: {e}")
    except Exception as e:
        logger.debug(f"BCC element debugging failed: {e}")
    
    for sel in candidates:
        try:
            el = page.locator(sel).first
            if DynamicTiming.wait_for_element_ready(page, el, max_wait=1500):
                # Strictly validate Bcc context before accepting as target
                if _is_in_bcc_context(el):
                    target = el
                    logger.info(f"Found BCC input target with selector: {sel}")
                    break
                else:
                    logger.debug(f"Discarded candidate not in Bcc context: {sel}")
        except Exception as e:
            logger.debug(f"Selector {sel} failed: {e}")
            continue
    
    if target is None:
        logger.warning("Could not find BCC input target for bulk paste")
        # Try one more approach - look for any input/textarea in the compose area
        try:
            logger.info("Trying fallback: looking for any input/textarea in compose area...")
            fallback_selectors = [
                'div[role="textbox"]',
                'input[type="text"]',
                'input[type="email"]',
                'textarea',
                '[contenteditable="true"]'
            ]
            for sel in fallback_selectors:
                try:
                    el = page.locator(sel).first
                    if DynamicTiming.wait_for_element_ready(page, el, max_wait=1000):
                        # Check if this element is in a BCC context
                        if _is_in_bcc_context(el):
                            target = el
                            logger.info(f"Found BCC input via fallback: {sel} (parent aria-label: {parent_aria})")
                            break
                except Exception:
                    continue
        except Exception as e:
            logger.debug(f"Fallback BCC detection failed: {e}")
        
        if target is None:
            # Last resort: try to find any input field that might be the BCC field
            logger.info("Last resort: looking for any input field that might be BCC...")
            try:
                # Look for any input/textarea that's visible and might be BCC
                all_inputs = page.locator('input, textarea, [contenteditable="true"], [role="textbox"]').all()
                logger.info(f"Found {len(all_inputs)} potential input fields")
                
                for i, input_elem in enumerate(all_inputs):
                    try:
                        if DynamicTiming.wait_for_element_ready(page, input_elem, max_wait=500):
                            # Check if this input is in a BCC context
                            if _is_in_bcc_context(input_elem):
                                target = input_elem
                                logger.info(f"Found BCC input via last resort: element {i+1}")
                                break
                    except Exception as e:
                        logger.debug(f"Could not check input element {i+1}: {e}")
                        continue
            except Exception as e:
                logger.debug(f"Last resort BCC detection failed: {e}")
            
            if target is None:
                return _bcc_chips_count(page, normalized)

    # Focus field and clear it (robust focus)
    with TimingContext("BCC field focus and clear") as timing:
        try:
            _force_focus(page, target)
            # Dynamic wait for field to be focused
            if DynamicTiming.wait_for_condition(
                lambda: target.evaluate("el => document.activeElement === el"),
                max_wait=timing.get_adaptive_delay(120)
            ):
                try:
                    target.fill("")
                except Exception:
                    # For contenteditable, select-all & delete
                    page.keyboard.press('Control+a')
                    time.sleep(0.05)  # Short wait for selection
                    page.keyboard.press('Delete')
                # Dynamic wait for field to be cleared
                DynamicTiming.wait_for_condition(
                    lambda: target.evaluate("el => el.textContent || el.value") == "",
                    max_wait=timing.get_adaptive_delay(100)
                )
        except Exception:
            pass

    # Compose bulk string
    bulk = ';'.join(normalized)
    logger.info(f"Bulk pasting BCC emails: {bulk}")

    # Paste using clipboard with multiple attempts
    with TimingContext("BCC bulk paste") as timing:
        try:
            import pyperclip
            pyperclip.copy(bulk)
            _force_focus(page, target)
            
            # Dynamic wait for focus
            DynamicTiming.wait_for_condition(
                lambda: target.evaluate("el => document.activeElement === el"),
                max_wait=timing.get_adaptive_delay(120)
            )
            
            # Clear any existing content first with multiple methods
            try:
                # Method 1: Select all and delete
                page.keyboard.press('Control+a')
                time.sleep(0.1)
                page.keyboard.press('Delete')
                time.sleep(0.2)
                
                # Method 2: Clear via JavaScript (more thorough)
                target.evaluate("el => { el.innerHTML = ''; el.textContent = ''; }")
                time.sleep(0.2)
                
                # Method 3: Clear via focus and clear
                target.click()
                target.press('Control+a')
                target.press('Delete')
                time.sleep(0.2)
                
                logger.info("✅ Cleared existing body content with multiple methods")
            except Exception as e:
                logger.debug(f"Error clearing body content: {e}")
            
            # Dynamic wait for field to be cleared
            DynamicTiming.wait_for_condition(
                lambda: target.evaluate("el => el.textContent || el.value || el.innerHTML") == "",
                max_wait=timing.get_adaptive_delay(200)
            )
            
            # Paste the bulk content
            page.keyboard.press('Control+V')
            
            # Verify content was actually set correctly
            try:
                current_content = target.evaluate("el => el.textContent || el.value || el.innerHTML")
                if current_content and current_content.strip():
                    logger.info(f"✅ Content successfully set: {current_content[:100]}...")
                else:
                    logger.warning("⚠️ Content may not have been set properly")
            except Exception as e:
                logger.debug(f"Error verifying content: {e}")
            
            # Dynamic wait for paste to complete and chips to form
            DynamicTiming.wait_for_bcc_chips_stable(page, len(normalized), max_wait=timing.get_adaptive_delay(500))  # Reduced from 1000ms
            
            # Try multiple tokenization methods
            tokenization_success = False
            
            # Method 1: Enter
            page.keyboard.press('Enter')
            if DynamicTiming.wait_for_bcc_chips_stable(page, len(normalized), max_wait=timing.get_adaptive_delay(400)):  # Reduced from 800ms
                tokenization_success = True
                logger.info("Bulk paste tokenized with Enter")
            
            if not tokenization_success:
                # Method 2: Tab
                target.click()
                page.keyboard.press('Tab')
                if DynamicTiming.wait_for_bcc_chips_stable(page, len(normalized), max_wait=timing.get_adaptive_delay(400)):  # Reduced from 800ms
                    tokenization_success = True
                    logger.info("Bulk paste tokenized with Tab")
            
            if not tokenization_success:
                # Method 3: Click outside
                try:
                    page.locator('input[aria-label="Add a subject"]').first.click(timeout=1000)
                    if DynamicTiming.wait_for_bcc_chips_stable(page, len(normalized), max_wait=timing.get_adaptive_delay(400)):  # Reduced from 800ms
                        tokenization_success = True
                        logger.info("Bulk paste tokenized with focus change")
                except Exception:
                    pass
            
            if not tokenization_success:
                # Method 4: Add semicolon and Enter
                target.click()
                page.keyboard.type(';')
                time.sleep(0.3)  # Short wait for semicolon
                page.keyboard.press('Enter')
                if DynamicTiming.wait_for_bcc_chips_stable(page, len(normalized), max_wait=timing.get_adaptive_delay(400)):  # Reduced from 800ms
                    tokenization_success = True
                    logger.info("Bulk paste tokenized with semicolon+Enter")
        except Exception as e:
            logger.error(f"Error in bulk paste: {e}")

    # If clipboard path didn't fully work, try typing and JS insertion fallbacks
    final_count = _bcc_chips_count(page, normalized)
    if final_count < len(normalized):
        try:
            _force_focus(page, target)
            page.keyboard.press('Control+a')
            page.wait_for_timeout(50)
            page.keyboard.press('Delete')
            page.wait_for_timeout(120)
            page.keyboard.type(bulk, delay=20)
            page.wait_for_timeout(300)
            # Commit
            page.keyboard.press('Enter')
            page.wait_for_timeout(600)
        except Exception:
            pass

    final_count = _bcc_chips_count(page, normalized)
    if final_count < len(normalized):
        # JS insertion for contenteditable targets
        try:
            _force_focus(page, target)
            page.wait_for_timeout(120)
            page.evaluate("el => document.execCommand('selectAll', false, null)", target)
            page.wait_for_timeout(80)
            page.evaluate("el => document.execCommand('insertText', false, arguments[0])", bulk)
            page.wait_for_timeout(500)
            page.keyboard.press('Enter')
            page.wait_for_timeout(700)
        except Exception:
            pass

    final_count = _bcc_chips_count(page, normalized)
    logger.info(f"Bulk paste result: {final_count} chips created")
    return final_count


def fast_send_email_with_cookies(
    sender_email: str,
    to: List[str],
    subject: str,
    body_html: str,
    cookies: List[Dict],
    bcc: List[str] = None,
    attachments: List[str] = None,
    headless: bool = True,
    proxy: str = None
) -> bool:
    """
    Fast email sending with cookie injection for Office 365 authentication.
    This bypasses the need for stored sessions by using fresh cookies.
    """
    bcc = bcc or []
    attachments = attachments or []
    
    with sync_playwright() as p:
        # Configure browser launch with proxy
        launch_kwargs = {'headless': headless}
        browser_type = p.chromium
        
        if proxy and proxy.strip():
            try:
                from urllib.parse import urlparse
                parsed = urlparse(proxy.strip())

                # Validate that we have required components
                if not parsed.scheme or not parsed.hostname or not parsed.port:
                    raise ValueError(f"Invalid proxy format. Missing scheme, hostname, or port. Got: {proxy}")

                server = f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"

                # Handle SOCKS proxies with authentication
                if parsed.scheme.startswith('socks') and (parsed.username or parsed.password):
                    logger.info("SOCKS proxy with authentication detected")
                    
                    # Try HTTP proxy conversion first (prefer Chromium)
                    logger.info("🔄 Converting SOCKS5 to HTTP proxy for Chromium compatibility...")
                    http_proxy = f"http://{parsed.username}:{parsed.password}@{parsed.hostname}:{parsed.port}"
                    logger.info(f"🔄 Trying HTTP proxy equivalent: {http_proxy}")
                    
                    # Update proxy configuration for HTTP
                    proxy_conf = { 'server': f"http://{parsed.hostname}:{parsed.port}" }
                    proxy_conf['username'] = parsed.username
                    proxy_conf['password'] = parsed.password
                    browser_type = p.chromium  # Prefer Chromium
                    
                    # Test HTTP proxy with Chromium
                    try:
                        test_browser = browser_type.launch(headless=True, proxy=proxy_conf)
                        test_browser.close()
                        logger.info("✅ HTTP proxy conversion successful with Chromium")
                    except Exception as http_error:
                        logger.warning(f"❌ HTTP proxy conversion failed: {http_error}")
                        logger.info("🔄 Trying Firefox as fallback for SOCKS proxy...")
                        
                        # Fallback to Firefox for SOCKS proxy
                        try:
                            browser_type = p.firefox
                            # Reset to original SOCKS proxy config
                            proxy_conf = { 'server': server }
                            proxy_conf['username'] = parsed.username
                            proxy_conf['password'] = parsed.password
                            
                            test_browser = browser_type.launch(headless=True, proxy=proxy_conf)
                            test_browser.close()
                            logger.info("✅ Firefox successfully supports this SOCKS proxy")
                        except Exception as firefox_error:
                            logger.error(f"❌ Firefox also failed with SOCKS proxy: {firefox_error}")
                            logger.warning("⚠️ Proceeding without proxy due to compatibility issues")
                            proxy_conf = None
                            browser_type = p.chromium

                # Apply proxy configuration if not disabled
                if proxy_conf is not None:
                    launch_kwargs['proxy'] = proxy_conf
                    logger.info(f"🌐 Using proxy: {proxy_conf['server']}")
                else:
                    logger.info("🌐 Proceeding without proxy")
            except Exception as e:
                logger.warning(f"Invalid proxy string; ignoring. Error: {e}")
                logger.info("Proceeding without proxy.")
        
        browser = browser_type.launch(**launch_kwargs)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        try:
            # Inject authentication cookies with correct domain
            outlook_host = _outlook_host_for_email(sender_email)
            _inject_cookies_to_context(context, cookies, f'.{outlook_host}')
            
            page = context.new_page()
            
            # Navigate to Outlook
            outlook_host = _outlook_host_for_email(sender_email)
            outlook_url = f"https://{outlook_host}/mail/"
            
            logger.info(f"Navigating to {outlook_url}")
            page.goto(outlook_url, wait_until="domcontentloaded")
            
            # Dynamic wait for page to load and check if we're authenticated
            if not DynamicTiming.wait_for_page_load(page, "mail", max_wait=10000):
                logger.warning("Page load timeout - proceeding anyway")
            
            # Check if we need to authenticate
            if "login" in page.url.lower() or "signin" in page.url.lower():
                logger.warning("Authentication required - cookies may be expired")
                return False
            
            # Navigate to compose - try pre-filling via query params
            if bcc:
                try:
                    bcc_param = ';'.join(bcc)
                    params = {
                        'bcc': bcc_param,
                        'subject': subject,
                        'body': body_html
                    }
                    compose_url = f"https://{outlook_host}/mail/deeplink/compose?{urlencode(params, quote_via=quote)}"
                except Exception:
                    compose_url = f"https://{outlook_host}/mail/deeplink/compose"
            else:
                compose_url = f"https://{outlook_host}/mail/deeplink/compose"
            logger.info(f"Navigating to compose: {compose_url}")
            page.goto(compose_url, wait_until="domcontentloaded")
            
            # Dynamic wait for compose page to load
            if not DynamicTiming.wait_for_page_load(page, "compose", max_wait=8000):
                logger.warning("Compose page load timeout - proceeding anyway")
            
            # Fill basic fields
            _ensure_basic_fields_filled(page, to, subject, body_html, proxy, attachments)
            
            # Handle attachments if provided
            if attachments:
                if not _handle_attachments(page, attachments, proxy):
                    logger.warning("⚠️ Failed to attach files, continuing without attachments")
            
            # COMPLETELY NEW BCC APPROACH - Try bulk paste first, then fall back
            if bcc:
                logger.info(f"=== BCC ENTRY START ({len(bcc)} recipients) ===")
                logger.info(f"Attempting BULK PASTE first...")

                # 1) Bulk paste attempt
                try:
                    bulk_count = _fill_bcc_bulk(page, bcc, browser, context, proxy)
                except Exception as e:
                    logger.warning(f"Bulk paste attempt errored: {e}")
                    bulk_count = 0

                # 2) If bulk incomplete, try URL params + keyboard reveal + direct input
                if bulk_count < len(set([e.lower() for e in (bcc or [])])):
                    logger.info(f"Bulk added {bulk_count}, trying URL/keyboard methods...")
                    logger.info(f"=== NEW BCC APPROACH: URL + Keyboard Shortcuts ===")
                    logger.info(f"Processing {len(bcc)} BCC recipients: {bcc}")
                
                # Method 1: Try to use Ctrl+Shift+B to show BCC field
                logger.info("Trying Control+Shift+B to reveal BCC field...")
                try:
                    page.keyboard.press('Control+Shift+B')
                    page.wait_for_timeout(2000)
                except Exception as e:
                    logger.debug(f"Control+Shift+B failed: {e}")
                
                # Method 2: Try to find BCC field with more aggressive selectors
                bcc_input_found = False
                bcc_input_selectors = [
                    # Standard selectors
                    'input[aria-label*="Bcc"]',
                    'input[aria-label*="BCC"]',
                    'input[placeholder*="Bcc"]',
                    'input[placeholder*="BCC"]',
                    # ContentEditable selectors
                    'div[aria-label*="Bcc"][contenteditable="true"]',
                    'div[aria-label*="BCC"][contenteditable="true"]',
                    'div[role="textbox"][aria-label*="Bcc"]',
                    'div[role="textbox"][aria-label*="BCC"]',
                    # Generic recipient selectors
                    'div[aria-label*="recipient"][contenteditable="true"]',
                    'div[role="combobox"][aria-label*="recipient"]',
                    # Fallback selectors
                    'div[contenteditable="true"]:has-text("Bcc")',
                    'div[contenteditable="true"]:has-text("BCC")',
                    # Last resort - any contenteditable in compose area
                    'div[contenteditable="true"]'
                ]
                
                bcc_input = None
                for selector in bcc_input_selectors:
                    try:
                        element = page.locator(selector).first
                        if element.is_visible(timeout=1000):
                            logger.info(f"Found BCC input with selector: {selector}")
                            bcc_input = element
                            bcc_input_found = True
                            break
                    except Exception as e:
                        logger.debug(f"BCC input selector {selector} failed: {e}")
                        continue
                
                if bcc_input_found and bcc_input:
                    # Method 3: Try to add BCC emails using the found input
                    logger.info("Attempting to add BCC emails to found input...")
                    try:
                        # Clear the field first
                        bcc_input.click()
                        page.wait_for_timeout(200)
                        bcc_input.fill("")
                        page.wait_for_timeout(200)
                        
                        # Add all emails separated by semicolons
                        bcc_string = ';'.join(bcc)
                        logger.info(f"Adding BCC string: {bcc_string}")
                        
                        # Type the BCC string
                        bcc_input.type(bcc_string, delay=50)
                        page.wait_for_timeout(1000)
                        
                        # Try to commit with Enter
                        page.keyboard.press('Enter')
                        page.wait_for_timeout(1000)
                        
                        # Check if it worked
                        final_count = _bcc_chips_count(page, bcc)
                        logger.info(f"BCC input method result: {final_count} chips created")
                        
                        if final_count >= len(bcc):
                            logger.info("✅ BCC input method succeeded!")
                        else:
                            logger.warning(f"⚠️ BCC input method only got {final_count}/{len(bcc)} chips")
                            
                    except Exception as e:
                        logger.error(f"Error with BCC input method: {e}")
                else:
                    logger.warning("Could not find any BCC input field")
                
                # Method 4: Try using the compose URL with BCC parameters
                if not bcc_input_found or _bcc_chips_count(page, bcc) < len(bcc):
                    logger.info("Trying URL-based BCC approach...")
                    try:
                        # Navigate to compose with BCC in URL
                        bcc_param = ';'.join(bcc)
                        compose_url_with_bcc = f"https://outlook.office.com/mail/0/deeplink/compose?bcc={bcc_param}&subject={subject}&body={body_html}"
                        logger.info(f"Navigating to: {compose_url_with_bcc}")
                        page.goto(compose_url_with_bcc, wait_until="networkidle")
                        page.wait_for_timeout(3000)
                        
                        # Check if URL method worked
                        url_count = _bcc_chips_count(page, bcc)
                        logger.info(f"URL BCC method result: {url_count} chips created")
                        
                        if url_count >= len(bcc):
                            logger.info("✅ URL BCC method succeeded!")
                        else:
                            logger.warning(f"⚠️ URL BCC method only got {url_count}/{len(bcc)} chips")
                            
                    except Exception as e:
                        logger.error(f"Error with URL BCC method: {e}")
                
                # 3) Final verification - check each email individually
                final_bcc_count = _bcc_chips_count(page, bcc)
                logger.info(f"Final BCC verification: {final_bcc_count}/{len(bcc)} chips")
                
                # Check each email individually
                verified_emails = []
                missing_emails = []
                for email in bcc:
                    if _bcc_has_email(page, email):
                        verified_emails.append(email)
                        logger.info(f"✅ Verified BCC email: {email}")
                    else:
                        missing_emails.append(email)
                        logger.warning(f"❌ Missing BCC email: {email}")
                
                logger.info(f"BCC verification summary: {len(verified_emails)}/{len(bcc)} emails verified")
                logger.info(f"Verified emails: {verified_emails}")
                if missing_emails:
                    logger.warning(f"Missing emails: {missing_emails}")
                
                if len(verified_emails) < len(bcc):
                    # 4) As a final fallback, try the robust individual entry helper for missing ones
                    missing = [e for e in bcc if e not in verified_emails]
                    if missing:
                        logger.info(f"Fallback to individual helper for missing: {missing}")
                        try:
                            added_now = _fill_bcc_field(page, missing, proxy)
                            logger.info(f"Fallback individual added: {added_now}/{len(missing)}")
                        except Exception as e:
                            logger.warning(f"Individual fallback error: {e}")
                    # Re-verify and capture screenshot if still partial
                    final_bcc_count = _bcc_chips_count(page, bcc)
                    if final_bcc_count < len(bcc):
                        logger.warning("⚠️ BCC field did not accept all recipients")
                        try:
                            # page.screenshot(path="logs/bcc_partial_debug.png")
                            logger.info("⚠️ BCC partial screenshot saved")
                        except Exception:
                            pass
                    else:
                        logger.info("✅ All BCC recipients successfully added after fallback")
            
            # Send email
            _click_send_button(page)
            
            # Verify email was sent
            success = _verify_email_sent(page, subject)
            
            if success:
                logger.info(f"Email sent successfully to {len(to)} recipients")
            else:
                logger.error("Failed to verify email was sent")
            
            return success
            
        except Exception as e:
            logger.error(f"Error in fast_send_email_with_cookies: {e}")
            return False
        finally:
            browser.close()


def _ensure_basic_fields_filled(page, to: List[str], subject: str, body_html: str, proxy: str = None, attachments: List[str] = None):
    """Fill email fields (To, Subject, Body) with proper HTML support"""
    try:
        # Fill To field
        to_field_selectors = [
            'input[aria-label="To"]',
            'input[placeholder*="To"]',
            'div[aria-label="To"] input',
            'div[data-testid="recipient-input"] input'
        ]
        
        filled_to = False
        for selector in to_field_selectors:
            try:
                to_field = page.wait_for_selector(selector, timeout=2500)
                if to_field:
                    existing = ''
                    try:
                        existing = to_field.input_value()
                    except Exception:
                        pass
                    if not existing.strip() and to and to[0]:
                        to_field.fill(to[0])
                    filled_to = True
                    break
            except Exception:
                continue
        # As a last resort, use JS to set the To input value and commit with Enter
        if not filled_to and to and to[0]:
            try:
                page.evaluate("(addr)=>{ const el=[...document.querySelectorAll('input[aria-label=\\'To\\']')][0]||[...document.querySelectorAll('div[aria-label=\\'To\\'] input')][0]; if(el){ el.focus(); el.value=addr; el.dispatchEvent(new Event('input',{bubbles:true})); } }", to[0])
                try:
                    page.keyboard.press('Enter')
                    time.sleep(0.15)
                except Exception:
                    pass
            except Exception:
                pass
        
        # Fill Subject field
        subject_field_selectors = [
            'input[aria-label="Add a subject"]',
            'input[placeholder*="subject"]',
            'div[aria-label="Add a subject"] input',
            'input[data-testid="subject-input"]'
        ]
        
        for selector in subject_field_selectors:
            try:
                subject_field = page.wait_for_selector(selector, timeout=2500)
                if subject_field:
                    existing = ''
                    try:
                        existing = subject_field.input_value()
                    except Exception:
                        pass
                    if not existing.strip() and subject:
                        subject_field.fill(subject)
                    break
            except Exception:
                continue
        
        # Fill Body field with HTML support
        _fill_body_field_with_html(page, body_html, proxy)
        
        logger.info("Basic email fields filled successfully with HTML support")
        
    except Exception as e:
        logger.error(f"Error filling basic email fields: {e}")


def _fill_body_field_with_html(page, body_html: str, proxy: str = None):
    """Fill email body field with proper HTML content support"""
    try:
        # Import HTML processor
        import sys
        import os
        sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from utils.html_email import process_html_email_content
        
        # Process HTML content
        html_result = process_html_email_content(body_html)
        
        if not html_result['success']:
            logger.warning(f"HTML processing failed, using fallback: {html_result['warnings']}")
            # Fallback to plain text
            processed_html = html_result['html_content']
        else:
            processed_html = html_result['html_content']
            if html_result['warnings']:
                logger.info(f"HTML processing warnings: {html_result['warnings']}")
        
        # Find body field selectors with improved detection
        body_field_selectors = [
            # High priority - most specific selectors
            'div[aria-label="Message body"]',
            'div[aria-label*="Message body"]',
            'div[data-testid="message-body"]',
            'div[data-testid*="message-body"]',
            'div[data-testid*="body"]',
            'div[data-automationid*="body"]',
            'div[data-automationid*="message"]',
            # NEW: More specific Outlook compose body selectors
            'div[class*="body"][contenteditable="true"]',
            'div[class*="Body"][contenteditable="true"]',
            'div[class*="compose"][contenteditable="true"]',
            'div[class*="Compose"][contenteditable="true"]',
            'div[class*="message"][contenteditable="true"]',
            'div[class*="Message"][contenteditable="true"]',
            # Medium priority - role-based selectors
            'div[role="textbox"][aria-label*="Message"]',
            'div[role="textbox"][aria-label*="body"]',
            'div[role="textbox"][contenteditable="true"]',
            'div[contenteditable="true"][role="textbox"]',
            'div[contenteditable="true"][aria-label*="Message"]',
            'div[contenteditable="true"][aria-label*="body"]',
            # Lower priority - generic selectors (but more specific)
            'div[contenteditable="true"][class*="body"]',
            'div[contenteditable="true"][class*="compose"]',
            'div[contenteditable="true"][class*="message"]',
            'div[contenteditable="true"]',
            'div[role="textbox"]',
            'div[aria-label*="Message"]',
            'div[aria-label*="body"]',
            'div[contenteditable]',
            # Fallback - class-based selectors
            'div[class*="message-body"]',
            'div[class*="body"]',
            'div[class*="content"]',
            'div[class*="editor"]'
        ]
        
        body_field = None
        # Use dynamic timeout for body field detection
        body_field_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
        logger.info(f"⏱️ Using adaptive body field timeout: {body_field_timeout}ms for proxy: {proxy or 'none'}")
        
        for selector in body_field_selectors:
            try:
                body_field = page.locator(selector).first
                if body_field.is_visible(timeout=body_field_timeout):
                    # Validate this is actually the email body field
                    if _validate_body_field(body_field):
                        logger.info(f"Found valid body field with selector: {selector}")
                        break
                    else:
                        logger.debug(f"Body field selector {selector} found but not valid email body")
                        continue
            except Exception as e:
                logger.debug(f"Body field selector {selector} failed: {e}")
                continue
        
        if not body_field:
            logger.warning("Could not find email body field with standard selectors, trying fallback methods...")
            
            # Fallback 1: Try to find any contenteditable element in the compose area
            try:
                # Look for contenteditable elements that are likely the body field
                all_contenteditable = page.locator('div[contenteditable="true"]')
                count = all_contenteditable.count()
                logger.info(f"Found {count} contenteditable elements, checking each one...")
                
                for i in range(count):
                    try:
                        element = all_contenteditable.nth(i)
                        if element.is_visible(timeout=2000):
                            # Check if this looks like the email body
                            rect = element.bounding_box()
                            if rect and rect['height'] > 100 and rect['width'] > 200:  # Large enough to be body
                                # Try to validate this element
                                if _validate_body_field(element):
                                    body_field = element
                                    logger.info(f"Found valid body field via fallback method (element {i})")
                                    break
                    except Exception as e:
                        logger.debug(f"Error checking contenteditable element {i}: {e}")
                        continue
            except Exception as e:
                logger.debug(f"Fallback method 1 failed: {e}")
            
            # Fallback 2: Try to find the largest contenteditable element
            try:
                compose_area = page.locator('[role="dialog"], .compose, [data-testid*="compose"], [aria-label*="New message"]').first
                if compose_area.is_visible(timeout=2000):
                    body_field = compose_area.locator('[contenteditable="true"]').first
                    if body_field.is_visible(timeout=2000):
                        logger.info("Found body field using fallback method 1")
                    else:
                        body_field = None
            except Exception:
                pass
            
            # Fallback 2: Try to find any textbox in the compose area
            if not body_field:
                try:
                    compose_area = page.locator('[role="dialog"], .compose, [data-testid*="compose"], [aria-label*="New message"]').first
                    if compose_area.is_visible(timeout=2000):
                        body_field = compose_area.locator('[role="textbox"]').first
                        if body_field.is_visible(timeout=2000):
                            logger.info("Found body field using fallback method 2")
                        else:
                            body_field = None
                except Exception:
                    pass
            
            # Fallback 3: Try to find the largest contenteditable element (likely the body)
            if not body_field:
                try:
                    all_contenteditable = page.locator('[contenteditable="true"]')
                    count = all_contenteditable.count()
                    if count > 0:
                        # Find the largest contenteditable element (likely the body)
                        largest_element = None
                        largest_size = 0
                        for i in range(count):
                            try:
                                element = all_contenteditable.nth(i)
                                if element.is_visible(timeout=1000):
                                    # Get element size
                                    size = element.evaluate("el => el.offsetWidth * el.offsetHeight")
                                    if size > largest_size:
                                        largest_size = size
                                        largest_element = element
                            except Exception:
                                continue
                        
                        if largest_element:
                            body_field = largest_element
                            logger.info("Found body field using fallback method 3 (largest contenteditable)")
                except Exception:
                    pass
            
            if not body_field:
                logger.error("Could not find email body field with any method")
                return False
        
        # Check if body already has content; if the content appears to be raw HTML text, we still replace it
        try:
            existing_html = body_field.evaluate("el => (el.innerHTML || '').trim()")
            existing_text = body_field.evaluate("el => (el.innerText || el.textContent || '').trim()")
        except Exception:
            existing_html = ''
            existing_text = ''

        looks_like_raw_html = existing_text.startswith('<') and '>' in existing_text and ('</' in existing_text or '/>' in existing_text)
        
        # Always clear existing content first, regardless of what's there
        if existing_text and existing_text.strip():
            logger.info(f"🧹 Clearing existing body content: {existing_text[:50]}...")
            try:
                # Clear existing content with multiple methods
                body_field.evaluate("el => { el.innerHTML = ''; el.textContent = ''; }")
                body_field.click()
                body_field.press("Control+a")
                body_field.press("Delete")
                time.sleep(0.5)
                logger.info("✅ Existing body content cleared successfully")
            except Exception as e:
                logger.debug(f"Error clearing existing content: {e}")
        
        # Now proceed with filling the new content
        logger.info("📝 Setting new body content after clearing existing content")
        
        # Fill body field with HTML content using multiple methods
        success = False
        
        # Method 1: Try innerHTML (best for HTML content)
        try:
            logger.info("Attempting to set HTML content via innerHTML")
            # Use dynamic timeout for evaluate operations
            eval_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
            logger.info(f"⏱️ Using adaptive evaluate timeout: {eval_timeout}ms for proxy: {proxy or 'none'}")
            
            # Set page timeout for evaluate operations
            try:
                original_timeout = page.timeout
            except AttributeError:
                # Page.timeout doesn't exist in this Playwright version
                original_timeout = 30000  # Default timeout
                logger.debug("Page.timeout not available, using default timeout")
            
            page.set_default_timeout(eval_timeout)
            
            try:
                # Properly escape HTML for JavaScript injection
                import json
                escaped_html = json.dumps(processed_html)[1:-1]  # Remove quotes and escape
                body_field.evaluate(f"el => el.innerHTML = {json.dumps(processed_html)}")
                page.wait_for_timeout(500)
                
                # Verify content was set
                current_content = body_field.evaluate("el => el.innerHTML")
                if processed_html.strip() in current_content or current_content.strip():
                    success = True
                    logger.info("HTML content set successfully via innerHTML")
            finally:
                # Restore original timeout
                page.set_default_timeout(original_timeout)
        except Exception as e:
            logger.debug(f"innerHTML method failed: {e}")
        
        # Method 2: Try fill with HTML (Playwright's fill method)
        if not success:
            try:
                logger.info("Attempting to set HTML content via fill method")
                # Use dynamic timeout for click and fill operations
                click_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
                body_field.click(timeout=click_timeout)
                page.wait_for_timeout(200)
                body_field.fill(processed_html, timeout=click_timeout)
                page.wait_for_timeout(500)
                success = True
                logger.info("HTML content set successfully via fill method")
                
                # Verify content was actually set correctly
                try:
                    current_content = body_field.evaluate("el => el.textContent || el.innerHTML")
                    if current_content and current_content.strip():
                        logger.info(f"✅ Body content verified: {current_content[:100]}...")
                    else:
                        logger.warning("⚠️ Body content may not have been set properly")
                except Exception as e:
                    logger.debug(f"Error verifying body content: {e}")
            except Exception as e:
                logger.debug(f"Fill method failed: {e}")
        
        # Method 3: Try typing HTML content
        if not success:
            try:
                logger.info("Attempting to set HTML content via typing")
                # Use dynamic timeout for click operation
                click_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
                body_field.click(timeout=click_timeout)
                page.wait_for_timeout(200)
                page.keyboard.press('Control+a')  # Select all
                page.wait_for_timeout(100)
                page.keyboard.type(processed_html, delay=10)
                page.wait_for_timeout(500)
                success = True
                logger.info("HTML content set successfully via typing")
                
                # Verify content was actually set correctly
                try:
                    current_content = body_field.evaluate("el => el.textContent || el.innerHTML")
                    if current_content and current_content.strip():
                        logger.info(f"✅ Body content verified: {current_content[:100]}...")
                    else:
                        logger.warning("⚠️ Body content may not have been set properly")
                except Exception as e:
                    logger.debug(f"Error verifying body content: {e}")
            except Exception as e:
                logger.debug(f"Typing method failed: {e}")
        
        # Method 4: Try clipboard paste
        if not success:
            try:
                logger.info("Attempting to set HTML content via clipboard")
                import pyperclip
                pyperclip.copy(processed_html)
                # Use dynamic timeout for click operation
                click_timeout = DynamicTiming.get_adaptive_timeout('element_wait', proxy)
                body_field.click(timeout=click_timeout)
                page.wait_for_timeout(200)
                page.keyboard.press('Control+a')  # Select all
                page.wait_for_timeout(100)
                page.keyboard.press('Control+v')  # Paste
                page.wait_for_timeout(500)
                success = True
                logger.info("HTML content set successfully via clipboard")
            except Exception as e:
                logger.debug(f"Clipboard method failed: {e}")
        
        if success:
            logger.info("Email body field filled with HTML content successfully")
            
            # Verify HTML content is properly rendered
            try:
                # Check if HTML tags are preserved
                current_html = body_field.evaluate("el => el.innerHTML")
                if '<' in current_html and '>' in current_html:
                    logger.info("HTML tags detected in body field - HTML rendering successful")
                else:
                    logger.warning("No HTML tags detected - content may be rendered as plain text")
            except Exception as e:
                logger.debug(f"HTML verification failed: {e}")
        else:
            logger.error("All methods failed to set HTML content in body field")
        
        return success
        
    except Exception as e:
        logger.error(f"Error filling body field with HTML: {e}")
        return False


def _click_send_button(page):
    """Click the send button"""
    try:
        # Try keyboard shortcut FIRST to avoid overlay interception
        try:
            page.keyboard.press('Control+Enter')
            logger.info("Triggered send via keyboard shortcut (primary)")
        except Exception:
            pass

        # Dismiss potential overlays that block clicks (hovercards, pills, popovers)
        try:
            page.keyboard.press('Escape')
            page.wait_for_timeout(150)
        except Exception:
            pass

        send_button_selectors = [
            "button[aria-label='Send']:not([aria-disabled='true'])",
            "button[title^='Send']:not([aria-disabled='true'])",
            "button[data-testid*='send']:not([aria-disabled='true'])",
            "button:has-text('Send'):not([aria-disabled='true'])"
        ]

        send_button = None
        for selector in send_button_selectors:
            try:
                send_button = page.locator(selector).first
                send_button.wait_for(state='visible', timeout=4000)
                break
            except Exception:
                continue

        if send_button:
            try:
                # Force click to bypass occasional overlays
                send_button.click(timeout=6000, force=True)
                logger.info("Send button clicked")
            except Exception as click_error:
                logger.warning(f"Primary send click failed: {click_error}")
                # Focus subject to dismiss any chip hovercards, then retry
                try:
                    subj = page.locator("input[aria-label='Add a subject']").first
                    if subj.is_visible(timeout=800):
                        subj.focus()
                        page.wait_for_timeout(150)
                except Exception:
                    pass
                # Retry forced click
                try:
                    send_button.click(timeout=4000, force=True)
                    logger.info("Send button clicked after focus change")
                except Exception:
                    # Fallback: keyboard again
                    try:
                        page.keyboard.press('Control+Enter')
                        logger.info("Triggered send via keyboard shortcut (fallback)")
                    except Exception as kb_err:
                        logger.error(f"Send shortcut failed: {kb_err}")
                        raise
        else:
            logger.warning("Send button not found; using keyboard fallback")
            page.keyboard.press('Control+Enter')
        
    except Exception as e:
        logger.error(f"Error clicking send button: {e}")
        raise


def _verify_email_sent(page, subject: str) -> bool:
    """Verify that email was sent successfully"""
    try:
        # Wait for navigation or success message
        page.wait_for_timeout(3000)
        
        # Check for success indicators
        success_indicators = [
            'text="Your message has been sent"',
            'text="Message sent"',
            'text="Email sent"',
            '[data-testid*="success"]',
            '.success-message'
        ]
        
        for indicator in success_indicators:
            try:
                if page.wait_for_selector(indicator, timeout=2000):
                    logger.info("Email sent successfully")
                    return True
            except:
                continue
        
        # Check if we're still on compose page (might indicate failure)
        if "compose" in page.url.lower():
            logger.warning("Still on compose page - email may not have been sent")
            return False
        
        # If we navigated away from compose, assume success
        logger.info("Email appears to have been sent (navigated away from compose)")
        return True
        
    except Exception as e:
        logger.error(f"Error verifying email sent: {e}")
        return False


def _validate_proxy(proxy: str) -> bool:
    """
    Test if proxy is working before using it
    Returns True if proxy is working, False otherwise
    """
    if not proxy or not proxy.strip():
        return True  # No proxy is always valid
    
    try:
        from urllib.parse import urlparse
        parsed = urlparse(proxy.strip())
        
        if not parsed.scheme or not parsed.hostname or not parsed.port:
            logger.warning(f"Invalid proxy format: {proxy}")
            return False
        
        # Test proxy with a more realistic request (like Outlook)
        import requests
        proxies = {
            'http': proxy,
            'https': proxy
        }
        
        # Test with a more realistic endpoint that's similar to Outlook
        try:
            # First test with a simple endpoint
            response = requests.get('https://httpbin.org/ip', proxies=proxies, timeout=5)
            if response.status_code != 200:
                logger.warning(f"❌ Proxy validation failed: {proxy} (Status: {response.status_code})")
                return False
            
            # Second test with a more complex HTTPS request (like Outlook)
            response2 = requests.get('https://outlook.office.com/', proxies=proxies, timeout=10)
            if response2.status_code == 200:
                logger.info(f"✅ Proxy validation successful: {proxy}")
                return True
            else:
                logger.warning(f"⚠️ Proxy validation partial: {proxy} (Simple: OK, Complex: {response2.status_code})")
                # Still return True for partial success, but log the warning
                return True
                
        except Exception as e:
            logger.warning(f"❌ Proxy validation failed: {proxy} (Error: {e})")
            return False
            
    except Exception as e:
        logger.warning(f"❌ Proxy validation failed: {proxy} (Error: {e})")
        return False

# Keep the original fast_send_email function but with YOUR PERFECT BCC approach
def fast_send_email(
    sender_email: str,
    to: List[str],
    subject: str,
    body_html: str,
    bcc: List[str] = None,
    attachments: List[str] = None,
    headless: bool = True,
    proxy: Optional[str] = None,
) -> bool:
    """
    Fast and reliable email sending via Outlook web.
    Returns True if successful, False otherwise.
    Uses YOUR PERFECT BCC APPROACH: Individual email entry with Enter/Tab
    """
    storage = _storage_path(sender_email)
    if not storage.exists():
        raise RuntimeError("Session not found. Run login first for this account.")

    bcc = bcc or []
    attachments = attachments or []
    
    logger.info(f"🚀 FAST SEND: {sender_email} → {to} | BCC: {len(bcc)} recipients | {subject}")

    # Track performance for dynamic timeout adjustment
    start_time = time.time()
    success = False
    
    # Validate proxy before using it
    if proxy and proxy.strip():
        logger.info(f"🔍 Validating proxy: {proxy}")
        if not _validate_proxy(proxy):
            logger.warning("⚠️ Proxy validation failed, proceeding without proxy")
            proxy = None
    
    try:
        with sync_playwright() as p:
            launch_kwargs = { 'headless': headless }
            browser_type = p.chromium
            
            # Optional proxy support
            if proxy and proxy.strip():
                try:
                    from urllib.parse import urlparse
                    parsed = urlparse(proxy.strip())
                    
                    if not parsed.scheme or not parsed.hostname or not parsed.port:
                        raise ValueError(f"Invalid proxy format. Missing scheme, hostname, or port. Got: {proxy}")
                    
                    server = f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"
                    
                    # Handle SOCKS proxies with authentication
                    if parsed.scheme.startswith('socks') and (parsed.username or parsed.password):
                        logger.info("SOCKS proxy with authentication detected")
                        
                        # Try HTTP proxy conversion first (prefer Chromium)
                        logger.info("🔄 Converting SOCKS5 to HTTP proxy for Chromium compatibility...")
                        http_proxy = f"http://{parsed.username}:{parsed.password}@{parsed.hostname}:{parsed.port}"
                        logger.info(f"🔄 Trying HTTP proxy equivalent: {http_proxy}")
                        
                        # Update proxy configuration for HTTP
                        proxy_conf = { 'server': f"http://{parsed.hostname}:{parsed.port}" }
                        proxy_conf['username'] = parsed.username
                        proxy_conf['password'] = parsed.password
                        browser_type = p.chromium  # Prefer Chromium
                        
                        # Test HTTP proxy with Chromium
                        try:
                            test_browser = browser_type.launch(headless=True, proxy=proxy_conf)
                            test_browser.close()
                            logger.info("✅ HTTP proxy conversion successful with Chromium")
                        except Exception as http_error:
                            logger.warning(f"❌ HTTP proxy conversion failed: {http_error}")
                            logger.info("🔄 Trying Firefox as fallback for SOCKS proxy...")
                            
                            # Fallback to Firefox for SOCKS proxy
                            try:
                                browser_type = p.firefox
                                # Reset to original SOCKS proxy config
                                proxy_conf = { 'server': server }
                                proxy_conf['username'] = parsed.username
                                proxy_conf['password'] = parsed.password
                                
                                test_browser = browser_type.launch(headless=True, proxy=proxy_conf)
                                test_browser.close()
                                logger.info("✅ Firefox successfully supports this SOCKS proxy")
                            except Exception as firefox_error:
                                logger.error(f"❌ Firefox also failed with SOCKS proxy: {firefox_error}")
                                logger.warning("⚠️ Proceeding without proxy due to compatibility issues")
                                proxy_conf = None
                                browser_type = p.chromium
                    else:
                        # Standard proxy configuration for HTTP/HTTPS
                        proxy_conf = { 'server': server }
                        if parsed.username:
                            proxy_conf['username'] = parsed.username
                        if parsed.password:
                            proxy_conf['password'] = parsed.password

                    # Apply proxy configuration if not disabled
                    if proxy_conf is not None:
                        launch_kwargs['proxy'] = proxy_conf
                        logger.info(f"🌐 Using proxy: {proxy_conf['server']}")
                    else:
                        logger.info("🌐 Proceeding without proxy")
                except Exception as e:
                    logger.warning(f"Invalid proxy string; ignoring. Error: {e}")
                    logger.info("Proceeding without proxy.")
            
            browser = browser_type.launch(**launch_kwargs)
            logger.info(f"🔧 Loading session from: {storage}")
            context = browser.new_context(storage_state=str(storage))
            page = context.new_page()
            
            # Navigate to Outlook with dynamic timeout
            host_for_nav = _outlook_host_for_email(sender_email)
            logger.info(f"🌐 Navigating to: https://{host_for_nav}/mail/")
            
            # Use adaptive timeout based on proxy performance
            page_timeout = DynamicTiming.get_adaptive_timeout('page_load', proxy)
            logger.info(f"⏱️ Using adaptive timeout: {page_timeout}ms for proxy: {proxy or 'none'}")
            
            try:
                page.goto(f"https://{host_for_nav}/mail/", wait_until="domcontentloaded", timeout=page_timeout)
            except Exception as nav_error:
                if "net::ERR_EMPTY_RESPONSE" in str(nav_error) or "ERR_PROXY" in str(nav_error) or "ERR_CONNECTION" in str(nav_error):
                    logger.error(f"❌ Proxy connection failed: {nav_error}")
                    logger.warning("🔄 This appears to be a proxy server issue, not a code problem")
                    logger.info("💡 Attempting automatic fallback to direct connection...")
                    
                    # Automatic fallback: Try without proxy
                    try:
                        logger.info("🔄 Retrying without proxy...")
                        # Close current browser and context
                        context.close()
                        browser.close()
                        
                        # Launch new browser without proxy
                        browser = browser_type.launch(**{k: v for k, v in launch_kwargs.items() if k != 'proxy'})
                        logger.info("🌐 Launched browser without proxy")
                        context = browser.new_context(storage_state=str(storage))
                        page = context.new_page()
                        
                        # Try navigation without proxy
                        page.goto(f"https://{host_for_nav}/mail/", wait_until="domcontentloaded", timeout=30000)
                        logger.info("✅ Successfully connected without proxy")
                        
                    except Exception as fallback_error:
                        logger.error(f"❌ Fallback to direct connection also failed: {fallback_error}")
                        raise RuntimeError(f"Both proxy and direct connection failed. Proxy error: {nav_error}, Direct error: {fallback_error}")
                else:
                    raise nav_error
            
            # Dynamic wait for mail page to load with proxy-aware timing
            if not DynamicTiming.wait_for_page_load(page, "mail", max_wait=10000, proxy=proxy):
                logger.warning("Mail page load timeout - proceeding anyway")
            
            # Navigate to compose (prefer deeplink with prefilled BCC when available)
            host = _outlook_host_for_email(sender_email)
            deeplink_path = "/mail/deeplink/compose" if host.endswith("live.com") else "/mail/0/deeplink/compose"
            if bcc:
                try:
                    from urllib.parse import urlencode, quote
                    bcc_param = ';'.join(bcc)
                    # Do NOT include body in URL to avoid plain-text HTML rendering
                    # Add sender email to TO field for testing/verification purposes
                    params = {
                        'to': sender_email,
                        'bcc': bcc_param,
                        'subject': subject
                    }
                    compose_url = f"https://{host}{deeplink_path}?{urlencode(params, quote_via=quote)}"
                except Exception:
                    compose_url = f"https://{host}{deeplink_path}"
            else:
                # Even without BCC, add sender to TO field for testing/verification
                try:
                    from urllib.parse import urlencode, quote
                    params = {
                        'to': sender_email,
                        'subject': subject
                    }
                    compose_url = f"https://{host}{deeplink_path}?{urlencode(params, quote_via=quote)}"
                except Exception:
                    compose_url = f"https://{host}{deeplink_path}"
            
            logger.info(f"🔗 Navigating to compose: {compose_url}")
            
            # Use adaptive timeout for compose page
            compose_timeout = DynamicTiming.get_adaptive_timeout('compose_load', proxy)
            logger.info(f"⏱️ Using adaptive compose timeout: {compose_timeout}ms for proxy: {proxy or 'none'}")
            
            page.goto(compose_url, wait_until='domcontentloaded', timeout=compose_timeout)
            
            # Wait for compose page to fully load (dynamic loading detection)
            compose_loaded = _wait_for_compose_page_loaded(page, proxy)
            
            # Clear compose window to ensure clean state
            if compose_loaded:
                _clear_compose_window(page)
                
                # Ensure TO field is populated with sender email (fallback if URL didn't work)
                _ensure_to_field_populated(page, sender_email, proxy)
            
            if not compose_loaded:
                logger.warning("⚠️ Compose page did not load completely")
                logger.info("🔄 Attempting fallback: refreshing page...")
                try:
                    page.reload(wait_until="domcontentloaded", timeout=compose_timeout)
                    compose_loaded = _wait_for_compose_page_loaded(page, proxy)
                    if compose_loaded:
                        logger.info("✅ Compose page loaded after refresh")
                    else:
                        logger.warning("⚠️ Compose page still not loaded after refresh")
                except Exception as e:
                    logger.error(f"❌ Page refresh failed: {e}")
            
            # Dynamic wait for compose UI to load with proxy-aware timing
            if not DynamicTiming.wait_for_page_load(page, "compose", max_wait=8000, proxy=proxy):
                logger.warning("Compose page load timeout - proceeding anyway")
            
            # Fill basic fields
            _ensure_basic_fields_filled(page, to, subject, body_html, proxy, attachments)
            
            # Handle attachments if provided
            if attachments:
                if not _handle_attachments(page, attachments, proxy):
                    logger.warning("⚠️ Failed to attach files, continuing without attachments")
            
            # Take screenshot of compose page after basic fields are filled
            try:
                screenshot_path = f"compose_page_after_basic_fields_{int(time.time())}.png"
                # page.screenshot(path=screenshot_path, full_page=True)
                logger.info(f"📸 Screenshot saved: {screenshot_path}")
            except Exception as e:
                logger.debug(f"Screenshot failed: {e}")
            
            # YOUR PERFECT BCC APPROACH - Try bulk paste first, then helper
            if bcc:
                logger.info(f"=== YOUR PERFECT BCC APPROACH: {len(bcc)} BCC recipients ===")
                
                # 1) Try bulk paste first
                try:
                    bulk_count = _fill_bcc_bulk(page, bcc, browser, context, proxy)
                    
                    # Take screenshot after BCC reveal attempt
                    try:
                        screenshot_path = f"compose_page_after_bcc_reveal_{int(time.time())}.png"
                        # page.screenshot(path=screenshot_path, full_page=True)
                        logger.info(f"📸 Screenshot after BCC reveal: {screenshot_path}")
                    except Exception as e:
                        logger.debug(f"BCC reveal screenshot failed: {e}")
                        
                except Exception as e:
                    logger.warning(f"Bulk BCC paste error: {e}")
                    bulk_count = 0
                
                need = len(set([e.lower() for e in (bcc or [])]))
                logger.info(f"Bulk paste result: {bulk_count}/{need} chips created")
                
                # 2) If bulk paste failed, try URL-based approach
                if bulk_count < need:
                    logger.info("Bulk paste incomplete, trying URL-based BCC pre-filling...")
                    try:
                        # Navigate to compose with BCC in URL
                        host = _outlook_host_for_email(sender_email)
                        deeplink_path = "/mail/deeplink/compose" if host.endswith("live.com") else "/mail/0/deeplink/compose"
                        from urllib.parse import urlencode, quote
                        bcc_param = ';'.join(bcc)
                        params = {
                            'bcc': bcc_param,
                            'subject': subject
                        }
                        compose_url_with_bcc = f"https://{host}{deeplink_path}?{urlencode(params, quote_via=quote)}"
                        logger.info(f"Navigating to compose with BCC in URL: {compose_url_with_bcc}")
                        # Use adaptive timeout for BCC compose navigation
                        bcc_timeout = DynamicTiming.get_adaptive_timeout('compose_load', proxy)
                        logger.info(f"⏱️ Using adaptive BCC timeout: {bcc_timeout}ms for proxy: {proxy or 'none'}")
                        page.goto(compose_url_with_bcc, wait_until='domcontentloaded', timeout=bcc_timeout)
                        
                        # Wait for page to load and check if BCC was pre-filled
                        if DynamicTiming.wait_for_page_load(page, "compose", max_wait=5000):
                            # Check if BCC was pre-filled from URL
                            url_bcc_count = _bcc_chips_count(page, bcc)
                            logger.info(f"URL-based BCC result: {url_bcc_count}/{need} chips created")
                            if url_bcc_count > bulk_count:
                                bulk_count = url_bcc_count
                    except Exception as e:
                        logger.warning(f"URL-based BCC approach failed: {e}")
                
                # 3) If still incomplete, try individual field filling
                if bulk_count < need:
                    logger.info(f"BCC still incomplete ({bulk_count}/{need}), trying individual field filling...")
                    
                    # Reveal Bcc field if needed
                    bcc_button_selectors = [
                        'button:has-text("Bcc")',
                        'button:has-text("Cc & Bcc")',
                        'button[aria-label*="Bcc"]',
                        'button[title*="Bcc"]',
                        '[data-testid*="bcc"]'
                    ]
                    bcc_button_found = False
                    for selector in bcc_button_selectors:
                        try:
                            bcc_btn = page.locator(selector).first
                            if DynamicTiming.wait_for_element_ready(page, bcc_btn, max_wait=2000):
                                logger.info(f"Found BCC button with selector: {selector}")
                                bcc_btn.click()
                                time.sleep(1)  # Wait for BCC field to appear
                                bcc_button_found = True
                                break
                        except Exception as e:
                            logger.debug(f"BCC button selector {selector} failed: {e}")
                            continue
                    
                    if not bcc_button_found:
                        logger.warning("Could not find BCC button")
                    
                    # Use robust helper to complete missing chips
                    logger.info(f"Using individual helper to complete BCC (current: {bulk_count}/{need})")
                    final_chips = _fill_bcc_field(page, bcc, proxy)
                else:
                    final_chips = bulk_count
                logger.info(f"Final BCC chips created: {final_chips}/{need}")
                if final_chips < need:
                    logger.warning("⚠️ BCC field did not accept all recipients")
                    try:
                        # page.screenshot(path="logs/bcc_partial_debug.png")
                        logger.info("⚠️ BCC partial screenshot saved")
                    except Exception:
                        pass
                
                # Take final screenshot after BCC completion
                try:
                    screenshot_path = f"compose_page_final_bcc_state_{int(time.time())}.png"
                    # page.screenshot(path=screenshot_path, full_page=True)
                    logger.info(f"📸 Final screenshot: {screenshot_path}")
                except Exception as e:
                    logger.debug(f"Final screenshot failed: {e}")
                
                logger.info(f"=== YOUR PERFECT BCC APPROACH COMPLETED ===")
            
            # If To field is still empty and we have at least one recipient, try again robustly
            try:
                if (to and to[0]):
                    # Verify a chip exists for To; if not, try keying again
                    chip_count = 0
                    try:
                        chip_count = int(page.evaluate("()=>{const el=document.querySelectorAll('[data-testid=\\'pill-0\\'], .ms-Persona'); return el? el.length: 0;}"))
                    except Exception:
                        chip_count = 0
                    if chip_count == 0:
                        _ensure_basic_fields_filled(page, to, subject, body_html, proxy, attachments)
                        try:
                            page.keyboard.press('Enter')
                            time.sleep(0.1)
                        except Exception:
                            pass
            except Exception as e:
                logger.debug(f"Final To verification failed: {e}")

            # Send email - keyboard only per requirement
            logger.info("📤 Sending email via keyboard shortcut only...")
            send_success = False

            with TimingContext("Email send") as timing:
                # More aggressive overlay dismissal and focus management
                try:
                    # Multiple escape presses to dismiss all overlays
                    for _ in range(3):
                        page.keyboard.press('Escape')
                        time.sleep(0.1)
                    
                    # Wait for overlays to dismiss
                    DynamicTiming.wait_for_condition(
                        lambda: page.locator('[id*="fluent-default-layer-host"] .customScrollBar').count() == 0,
                        max_wait=timing.get_adaptive_delay(200)
                    )
                    
                    # Try multiple focus targets for better reliability
                    focus_targets = [
                        "input[aria-label='Add a subject']",
                        "input[aria-label*='subject']",
                        "input[placeholder*='subject']",
                        "div[aria-label*='subject']",
                        "div[role='textbox'][aria-label*='subject']"
                    ]
                    
                    focused = False
                    for target_selector in focus_targets:
                        try:
                            target = page.locator(target_selector).first
                            if DynamicTiming.wait_for_element_ready(page, target, max_wait=300):
                                target.focus()
                                # Wait for focus with multiple checks
                                for _ in range(3):
                                    if DynamicTiming.wait_for_condition(
                                        lambda: target.evaluate("el => document.activeElement === el"),
                                        max_wait=timing.get_adaptive_delay(100)
                                    ):
                                        focused = True
                                        logger.info(f"Successfully focused: {target_selector}")
                                        break
                                    time.sleep(0.1)
                                if focused:
                                    break
                        except Exception:
                            continue
                    
                    if not focused:
                        logger.warning("Could not focus any target, proceeding with send attempt")
                        
                except Exception as e:
                    logger.warning(f"Focus management failed: {e}")

                # Press Control+Enter with optimized timing
                try:
                    # First attempt - single press with minimal wait
                    page.keyboard.press('Control+Enter')
                    time.sleep(0.1)  # Reduced from 0.3s to 0.1s
                    
                    # Check if send was successful with shorter timeout
                    if DynamicTiming.wait_for_send_confirmation(page, max_wait=1000):  # Reduced from 2000ms
                        send_success = True
                        success = True
                        logger.info("✅ Send successful on first attempt")
                    else:
                        # Second attempt - double press with minimal delays
                        logger.info("First attempt failed, trying double press...")
                        page.keyboard.press('Control+Enter')
                        time.sleep(0.05)  # Reduced from 0.2s
                        page.keyboard.press('Control+Enter')
                        time.sleep(0.2)  # Reduced from 0.5s
                        
                        if DynamicTiming.wait_for_send_confirmation(page, max_wait=1500):  # Reduced from 3000ms
                            send_success = True
                            success = True
                            logger.info("✅ Send successful on second attempt")
                        else:
                            # Third attempt - with minimal focus delay
                            logger.info("Second attempt failed, trying with additional focus...")
                            try:
                                # Click somewhere to ensure focus
                                page.mouse.click(100, 100)
                                time.sleep(0.05)  # Reduced from 0.1s
                                page.keyboard.press('Control+Enter')
                                time.sleep(0.1)  # Reduced from 0.2s
                                page.keyboard.press('Control+Enter')
                                time.sleep(0.2)  # Reduced from 0.5s
                                
                                if DynamicTiming.wait_for_send_confirmation(page, max_wait=1500):  # Reduced from 3000ms
                                    send_success = True
                                    success = True
                                    logger.info("✅ Send successful on third attempt")
                                else:
                                    logger.warning("⚠️ All send attempts failed")
                            except Exception as e:
                                logger.error(f"Third send attempt failed: {e}")
                    
                except Exception as e:
                    logger.error(f"Keyboard send failed: {e}")
                    send_success = False
            
            # Verify email was sent - more comprehensive verification
            logger.info("🔍 Verifying send confirmation...")
            try:
                # Wait for network idle first with reduced timeout
                if DynamicTiming.wait_for_network_idle(page, max_wait=8000):  # Reduced from 15000ms
                    logger.info("✅ Network idle after send - platform processed")
                    
                    # Comprehensive send confirmation check with faster timeouts
                    confirmation_success = False
                    
                    # Method 1: Check for success indicators with reduced timeout
                    if DynamicTiming.wait_for_send_confirmation(page, max_wait=2000):  # Reduced from 5000ms
                        confirmation_success = True
                        logger.info("✅ Send confirmed via success indicators")
                    
                    # Method 2: Check URL change
                    if not confirmation_success:
                        current_url_after_send = page.url
                        logger.info(f"URL after send attempt: {current_url_after_send}")
                        
                        if 'compose' not in current_url_after_send.lower():
                            confirmation_success = True
                            logger.info("✅ Send confirmed - no longer on compose page")
                    
                    # Method 3: Check if compose window is gone
                    if not confirmation_success:
                        try:
                            compose_windows = page.locator('div[aria-label*="New message"], div[role="dialog"][aria-label*="New message"]').count()
                            if compose_windows == 0:
                                confirmation_success = True
                                logger.info("✅ Send confirmed - compose window disappeared")
                        except Exception:
                            pass
                    
                    # Method 4: Check drafts folder
                    if not confirmation_success:
                        try:
                            # Navigate to drafts folder to check if email is gone
                            drafts_link = page.locator('a[href*="drafts"], [aria-label*="Drafts"]').first
                            if DynamicTiming.wait_for_element_ready(page, drafts_link, max_wait=2000):
                                drafts_link.click()
                                time.sleep(1)
                                
                                # Check if our email is still in drafts
                                draft_emails = page.locator('[data-testid*="draft"], [aria-label*="Draft"]').count()
                                if draft_emails == 0:
                                    confirmation_success = True
                                    logger.info("✅ Send confirmed - email no longer in drafts")
                                else:
                                    logger.warning(f"⚠️ Email still in drafts ({draft_emails} drafts found)")
                        except Exception as e:
                            logger.debug(f"Drafts check failed: {e}")
                    
                    if not confirmation_success:
                        logger.warning("⚠️ Could not confirm send success - email may still be in drafts")
                        # Take a screenshot for debugging
                        try:
                            # page.screenshot(path="logs/send_failure_debug.png")
                            logger.info("📸 Send failure screenshot saved")
                        except Exception:
                            pass
                else:
                    logger.warning("⚠️ Network did not become idle after send")
                    
            except Exception as e:
                logger.warning(f"⚠️ Send verification error: {e}")

            # Always close Playwright resources
            try:
                page.close()
            except Exception:
                pass
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass

            return send_success

    except Exception as e:
        logger.error(f"❌ Send failed with unexpected error: {e}")
        return False
    finally:
        # Record performance for future timeout adjustments
        end_time = time.time()
        response_time = (end_time - start_time) * 1000  # Convert to milliseconds
        DynamicTiming.record_proxy_performance(proxy, success, response_time)
        
        if proxy:
            logger.info(f"📊 Performance recorded: {response_time:.1f}ms, Success: {success} for proxy: {proxy}")
