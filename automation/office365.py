"""
Office365 browser automation using Playwright.
- Login and persist session per account (storage state file)
- Send email with BCC using saved session
"""

import os
from typing import List
from pathlib import Path
from playwright.sync_api import sync_playwright
from urllib.parse import urlencode, quote


SESSIONS_DIR = Path("sessions")
SESSIONS_DIR.mkdir(exist_ok=True)


def _storage_path(email: str) -> Path:
    safe = email.replace("@", "_at_").replace(".", "_")
    return SESSIONS_DIR / f"office365_{safe}.json"


def playwright_install_if_needed():
    """Ensure playwright browsers are installed."""
    # No-op here; documented for manual run: `playwright install`.
    pass


def login(email: str, password: str, headless: bool = True) -> str:
    """Log in to Office365 via browser and save storage state.
    Returns path to storage state file.
    """
    playwright_install_if_needed()
    storage = _storage_path(email)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        # Navigate to Outlook web login
        page.goto("https://outlook.office.com/mail/")
        page.wait_for_load_state('domcontentloaded')

        # Email entry (robust selectors)
        email_box = page.locator("input[name='loginfmt'], input[type='email']").first
        email_box.wait_for(state='visible', timeout=30000)
        email_box.fill(email)
        # Hitting Enter is more reliable than clicking the disabled Next
        email_box.press('Enter')
        # Or wait until the Next button becomes enabled then click as a fallback
        try:
            page.locator("#idSIButton9:not([disabled])").first.wait_for(timeout=5000)
            page.locator("#idSIButton9").click()
        except Exception:
            pass

        # Password entry
        pwd_box = page.locator("input[name='passwd'], input[type='password']").first
        pwd_box.wait_for(state='visible', timeout=30000)
        pwd_box.fill(password)

        # Sign in
        try:
            page.locator("#idSIButton9:not([disabled])").first.wait_for(timeout=5000)
            page.locator("#idSIButton9").click()
        except Exception:
            pwd_box.press('Enter')

        # Handle stay signed in prompt if appears
        try:
            page.locator("#idSIButton9").click(timeout=7000)
        except Exception:
            pass

        # Wait for mailbox
        try:
            page.wait_for_url("**/mail/*", timeout=90000)
        except Exception:
            # Fallback: wait for New mail button
            page.locator('[aria-label="New mail"], button:has-text("New mail")').first.wait_for(timeout=30000)

        # Persist session
        context.storage_state(path=str(storage))
        browser.close()
    return str(storage)


def send_with_bcc(
    sender_email: str,
    to: List[str],
    subject: str,
    body_html: str,
    bcc: List[str] = None,
    attachments: List[str] = None,
    headless: bool = True,
) -> None:
    """Fast and reliable email sending via Outlook web."""
    from loguru import logger
    
    storage = _storage_path(sender_email)
    if not storage.exists():
        raise RuntimeError("Session not found. Run login first for this account.")

    bcc = bcc or []
    attachments = attachments or []
    
    logger.info(f"🚀 Fast email send: {sender_email} → {to} | {subject}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(storage))
        page = context.new_page()
        
        # Direct compose URL with pre-filled data
        params = {
            'to': ";".join(to) if to else '',
            'bcc': ";".join(bcc) if bcc else '',
            'subject': subject or '',
            'body': body_html or ''
        }
        compose_url = "https://outlook.office.com/mail/deeplink/compose?" + urlencode(params, quote_via=quote)
        
        logger.info("📧 Opening compose window...")
        page.goto(compose_url)
        page.wait_for_load_state('networkidle', timeout=15000)

        # Ensure the compose editor is visible (try several patterns) and only fallback to manual entry if needed
        candidate_selectors = [
            "role=textbox[name='To']",
            "[aria-label='To']",
            "div[role='combobox'][aria-label='To']",
            "input[aria-label='To']",
            "[data-automationid='ToRecipientWell']",
        ]
        to_found = False
        for sel in candidate_selectors:
            try:
                page.locator(sel).first.wait_for(timeout=8000)
                to_found = True
                break
            except Exception:
                continue
        if not to_found:
            # Fallback: open a fresh compose window then re-check
            try:
                page.keyboard.press('Control+n')
            except Exception:
                pass
            page.wait_for_timeout(1000)
            for sel in candidate_selectors:
                try:
                    page.locator(sel).first.wait_for(timeout=8000)
                    to_found = True
                    break
                except Exception:
                    continue
        # If still not found, proceed; deeplink may have already populated and focused correctly

        # To
        if to:
            # Try multiple ways to target To field
            filled = False
            for sel in [
                "role=textbox[name='To']",
                "input[aria-label='To']",
                "div[role='combobox'][aria-label='To']",
                "[aria-label='To']",
            ]:
                try:
                    box = page.locator(sel).first
                    box.click()
                    # Some To fields are contenteditable and do not support fill
                    try:
                        box.fill(", ".join(to))
                    except Exception:
                        page.keyboard.type(", ".join(to))
                    filled = True
                    break
                except Exception:
                    continue
            if not filled:
                # Final fallback: focus body and use Ctrl+L to move? Instead, click first input on the page
                page.locator("input, div[contenteditable='true']").first.click()
                page.keyboard.type(", ".join(to))

        # BCC toggle and fill if present
        try:
            page.get_by_role("button", name="Bcc").click(timeout=3000)
        except Exception:
            # Some tenants show Bcc by default
            pass
        if bcc:
            try:
                bcc_box = page.get_by_role("textbox", name="Bcc").first
                bcc_box.click()
                bcc_box.fill(", ".join(bcc))
            except Exception:
                try:
                    bcc_alt = page.locator("input[aria-label='Bcc'], div[role='combobox'][aria-label='Bcc']").first
                    bcc_alt.click()
                    try:
                        bcc_alt.fill(", ".join(bcc))
                    except Exception:
                        page.keyboard.type(", ".join(bcc))
                except Exception:
                    pass

        # Subject - Multiple robust selectors
        if subject:
            subject_filled = False
            subject_selectors = [
                "input[placeholder='Add a subject']",
                "input[aria-label='Subject']", 
                "input[data-automationid='Subject']",
                "input[name='subject']",
                "input[type='text'][placeholder*='subject']",
                "input[type='text'][placeholder*='Subject']"
            ]
            
            for selector in subject_selectors:
                try:
                    subject_field = page.locator(selector).first
                    if subject_field.is_visible(timeout=5000):
                        subject_field.click()
                        subject_field.fill(subject)
                        subject_filled = True
                        break
                except Exception:
                    continue
            
            # Fallback: try role-based selector with shorter timeout
            if not subject_filled:
                try:
                    subject_field = page.get_by_role("textbox", name="Subject").first
                    if subject_field.is_visible(timeout=3000):
                        subject_field.click()
                        subject_field.fill(subject)
                        subject_filled = True
                except Exception:
                    pass
            
            # Final fallback: type directly
            if not subject_filled:
                try:
                    page.keyboard.press('Tab')  # Navigate to subject field
                    page.keyboard.type(subject)
                except Exception:
                    logger.warning("Could not fill subject field")

        # Body - Multiple robust approaches
        if body_html:
            logger.info("Attempting to fill email body...")
            # Small delay to ensure compose window is fully loaded
            page.wait_for_timeout(2000)
            body_filled = False
            
            # Method 1: Try iframe approach (most common in Outlook)
            try:
                iframe = page.frame_locator('iframe[title="Message body"]')
                body_editor = iframe.locator("div[contenteditable=true]")
                if body_editor.is_visible(timeout=5000):
                    body_editor.click()
                    body_editor.fill(body_html)
                    body_filled = True
                    logger.info("Body filled via iframe")
            except Exception as e:
                logger.warning(f"Iframe body fill failed: {e}")
            
            # Method 2: Try direct contenteditable div
            if not body_filled:
                try:
                    body_editor = page.locator('div[contenteditable="true"]').first
                    if body_editor.is_visible(timeout=5000):
                        body_editor.click()
                        body_editor.fill(body_html)
                        body_filled = True
                        logger.info("Body filled via direct contenteditable")
                except Exception as e:
                    logger.warning(f"Direct contenteditable failed: {e}")
            
            # Method 3: Try alternative iframe selectors
            if not body_filled:
                iframe_selectors = [
                    'iframe[title="Message body"]',
                    'iframe[title="Rich Text Editor"]',
                    'iframe[title="Message"]',
                    'iframe[data-automationid="MessageBody"]'
                ]
                for iframe_sel in iframe_selectors:
                    try:
                        iframe = page.frame_locator(iframe_sel)
                        body_editor = iframe.locator("div[contenteditable=true], body")
                        if body_editor.is_visible(timeout=3000):
                            body_editor.click()
                            body_editor.fill(body_html)
                            body_filled = True
                            logger.info(f"Body filled via iframe: {iframe_sel}")
                            break
                    except Exception:
                        continue
            
            # Method 4: Try role-based selector
            if not body_filled:
                try:
                    body_editor = page.get_by_role("textbox", name="Message body").first
                    if body_editor.is_visible(timeout=3000):
                        body_editor.click()
                        body_editor.fill(body_html)
                        body_filled = True
                        logger.info("Body filled via role selector")
                except Exception as e:
                    logger.warning(f"Role-based body fill failed: {e}")
            
            # Method 5: Final fallback - keyboard typing
            if not body_filled:
                try:
                    # Focus on the compose area and type directly
                    page.keyboard.press('Tab')  # Navigate to body area
                    page.keyboard.press('Tab')
                    page.keyboard.type(body_html)
                    body_filled = True
                    logger.info("Body filled via keyboard typing")
                except Exception as e:
                    logger.warning(f"Keyboard typing failed: {e}")
            
            if not body_filled:
                logger.error("Failed to fill email body - all methods failed")

        # Attachments
        for file_path in attachments:
            if os.path.isfile(file_path):
                page.get_by_label("Attach").click()
                page.set_input_files("input[type=file]", file_path)

        # Send - Multiple approaches
        logger.info("Attempting to send email...")
        send_success = False
        
        # Try primary send button
        try:
            send_btn = page.get_by_role("button", name="Send").first
            if send_btn.is_visible(timeout=5000):
                send_btn.click(timeout=10000)
                send_success = True
                logger.info("Email sent via primary Send button")
        except Exception as e:
            logger.warning(f"Primary send button failed: {e}")
        
        # Try alternative send button
        if not send_success:
            try:
                send_btn = page.locator("button:has-text('Send')").first
                if send_btn.is_visible(timeout=5000):
                    send_btn.click(timeout=10000)
                    send_success = True
                    logger.info("Email sent via alternative Send button")
            except Exception as e:
                logger.warning(f"Alternative send button failed: {e}")
        
        # Try keyboard shortcut
        if not send_success:
            try:
                page.keyboard.press('Control+Enter')
                send_success = True
                logger.info("Email sent via Ctrl+Enter")
            except Exception as e:
                logger.warning(f"Keyboard shortcut failed: {e}")

        # Wait for confirmation
        if send_success:
            try:
                page.wait_for_selector("text=Sent", timeout=10000)
                logger.info("Email send confirmed")
            except Exception:
                try:
                    page.wait_for_url("**/mail/**", timeout=10000)
                    logger.info("Email send confirmed via URL change")
                except Exception:
                    logger.warning("Could not confirm email send, but attempt was made")
        else:
            logger.error("Failed to send email - no method worked")
            
        browser.close()


