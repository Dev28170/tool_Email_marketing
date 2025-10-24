"""
Office 365 Cookie Management System
Handles cookie injection for Office 365 authentication
"""

import json
import base64
from typing import Dict, List, Optional
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

class Office365CookieManager:
    """Manages Office 365 authentication cookies for programmatic login"""
    
    def __init__(self):
        self.accounts = {}  # Store account data: {email: {cookies, metadata}}
    
    def add_account(self, email: str, cookies_data: str, account_type: str = "paid") -> bool:
        """
        Add an Office 365 account with cookie data
        
        Args:
            email: Office 365 email address
            cookies_data: JSON string containing cookie information
            account_type: "free" or "paid"
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Parse cookie data
            cookies = self._parse_cookie_data(cookies_data)
            
            if not cookies:
                logger.error(f"Failed to parse cookie data for {email}")
                return False
            
            # Validate cookies
            if not self._validate_cookies(cookies):
                logger.error(f"Invalid cookie data for {email}")
                return False
            
            # Store account data
            self.accounts[email] = {
                'cookies': cookies,
                'account_type': account_type,
                'added_at': datetime.now().isoformat(),
                'last_used': None,
                'status': 'active'
            }
            
            logger.info(f"Successfully added Office 365 account: {email}")
            return True
            
        except Exception as e:
            logger.error(f"Error adding account {email}: {str(e)}")
            return False
    
    def _parse_cookie_data(self, cookies_data: str) -> List[Dict]:
        """Parse cookie data from various formats"""
        try:
            text = (cookies_data or "").strip()
            # New simple triple-value format: VAL1###VAL2###VAL3
            # Interpreted as ESTSAUTHPERSISTENT, ESTSAUTH, ESTSAUTHLIGHT (in this order)
            if '###' in text and not text.startswith('['):
                parts = [p.strip() for p in text.split('###')]
                # Allow accidental extra separators; keep first three meaningful parts
                parts = [p for p in parts if p] 
                if len(parts) >= 3:
                    names = ['ESTSAUTHPERSISTENT', 'ESTSAUTH', 'ESTSAUTHLIGHT']
                    cookies: List[Dict] = []
                    for name, raw in zip(names, parts[:3]):
                        # Accept formats like NAME=VALUE or just VALUE
                        value = raw
                        if '=' in raw:
                            try:
                                # Take everything after first '=' as the value
                                value = raw.split('=', 1)[1]
                            except Exception:
                                value = raw
                        cookies.append({
                            'name': name,
                            'value': value,
                            'domain': '.login.microsoftonline.com',
                            'path': '/',
                            'secure': True,
                            'httpOnly': True,
                            'sameSite': 'None'
                        })
                    return cookies

            # Try to parse as JSON first
            if text.startswith('['):
                return json.loads(text)
            
            # Try to extract from JavaScript code
            if 'JSON.parse' in text:
                # Extract JSON from JavaScript
                start = text.find('JSON.parse(`') + 12
                end = text.find('`);', start)
                if start > 11 and end > start:
                    json_str = text[start:end]
                    return json.loads(json_str)
            
            # Try to parse as base64 encoded
            try:
                decoded = base64.b64decode(text).decode('utf-8')
                return json.loads(decoded)
            except:
                pass
            
            logger.error("Unable to parse cookie data format")
            return []
            
        except Exception as e:
            logger.error(f"Error parsing cookie data: {str(e)}")
            return []
    
    def _validate_cookies(self, cookies: List[Dict]) -> bool:
        """Validate that cookies contain required Office 365 authentication data"""
        if not cookies or not isinstance(cookies, list):
            return False
        
        # Check for required Office 365 authentication cookies
        required_cookies = ['ESTSAUTHPERSISTENT', 'ESTSAUTH', 'ESTSAUTHLIGHT']
        cookie_names = [cookie.get('name', '') for cookie in cookies]
        
        for required in required_cookies:
            if required not in cookie_names:
                logger.warning(f"Missing required cookie: {required}")
                return False
        
        return True
    
    def get_account(self, email: str) -> Optional[Dict]:
        """Get account data by email"""
        return self.accounts.get(email)
    
    def get_all_accounts(self) -> Dict:
        """Get all stored accounts"""
        return self.accounts.copy()
    
    def remove_account(self, email: str) -> bool:
        """Remove an account"""
        if email in self.accounts:
            del self.accounts[email]
            logger.info(f"Removed account: {email}")
            return True
        return False
    
    def update_account_status(self, email: str, status: str) -> bool:
        """Update account status (active, expired, error)"""
        if email in self.accounts:
            self.accounts[email]['status'] = status
            self.accounts[email]['last_used'] = datetime.now().isoformat()
            return True
        return False
    
    def get_cookies_for_injection(self, email: str) -> List[Dict]:
        """Get cookies formatted for browser injection"""
        account = self.get_account(email)
        if not account:
            return []
        
        cookies = account['cookies']
        formatted_cookies = []
        
        for cookie in cookies:
            # Convert to Playwright format
            formatted_cookie = {
                'name': cookie.get('name', ''),
                'value': cookie.get('value', ''),
                'domain': cookie.get('domain', '.login.microsoftonline.com'),
                'path': cookie.get('path', '/'),
                'secure': cookie.get('secure', True),
                'httpOnly': cookie.get('httpOnly', True),
                'sameSite': cookie.get('sameSite', 'none')
            }
            
            # Handle expiration
            if 'expirationDate' in cookie:
                exp_date = cookie['expirationDate']
                if isinstance(exp_date, (int, float)):
                    # Convert from milliseconds to seconds
                    exp_seconds = exp_date / 1000
                    formatted_cookie['expires'] = exp_seconds
            
            formatted_cookies.append(formatted_cookie)
        
        return formatted_cookies
    
    def is_cookie_valid(self, email: str) -> bool:
        """Check if cookies are still valid (not expired)"""
        account = self.get_account(email)
        if not account:
            return False
        
        cookies = account['cookies']
        current_time = datetime.now().timestamp() * 1000  # Convert to milliseconds
        
        for cookie in cookies:
            if 'expirationDate' in cookie:
                exp_date = cookie['expirationDate']
                if isinstance(exp_date, (int, float)) and exp_date < current_time:
                    logger.warning(f"Cookie {cookie.get('name')} expired for {email}")
                    return False
        
        return True
    
    def get_active_accounts(self) -> List[str]:
        """Get list of active account emails"""
        active_accounts = []
        for email, account in self.accounts.items():
            if account['status'] == 'active' and self.is_cookie_valid(email):
                active_accounts.append(email)
        return active_accounts

# Global instance
cookie_manager = Office365CookieManager()





