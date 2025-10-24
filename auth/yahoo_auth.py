"""
Yahoo Mail OAuth2.0 authentication implementation
Handles Yahoo Mail API authentication and email sending
"""

import asyncio
import aiohttp
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from .base_auth import BaseOAuth2Provider, OAuth2Error
from config import config

class YahooAuth(BaseOAuth2Provider):
    """Yahoo Mail OAuth2.0 authentication provider"""
    
    def __init__(self):
        super().__init__('yahoo')
    
    def _get_auth_params(self) -> Dict[str, str]:
        """Get Yahoo-specific authorization parameters"""
        return {}
    
    def _get_token_params(self) -> Dict[str, str]:
        """Get Yahoo-specific token request parameters"""
        return {}
    
    def _get_refresh_params(self) -> Dict[str, str]:
        """Get Yahoo-specific refresh token parameters"""
        return {}
    
    def _get_user_info_url(self) -> str:
        """Get Yahoo user info endpoint"""
        return "https://api.login.yahoo.com/openid/v1/userinfo"
    
    def _get_revoke_url(self) -> Optional[str]:
        """Get Yahoo token revocation endpoint"""
        return "https://api.login.yahoo.com/oauth2/revoke"
    
    async def get_user_info(self, access_token: str) -> Dict[str, Any]:
        """Get Yahoo user information"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        async with aiohttp.ClientSession() as session:
            async with session.get(self._get_user_info_url(), headers=headers) as response:
                if response.status != 200:
                    error_text = await response.text()
                    raise OAuth2Error(f"Failed to get user info: {response.status} - {error_text}")
                
                user_data = await response.json()
                
                # Standardize user info format
                return {
                    'id': user_data.get('sub'),
                    'email': user_data.get('email'),
                    'name': user_data.get('name'),
                    'first_name': user_data.get('given_name'),
                    'last_name': user_data.get('family_name'),
                    'provider': 'yahoo',
                    'verified_email': user_data.get('email_verified', False)
                }
    
    def _create_message(self, sender: str, to: List[str], subject: str, 
                       body_html: str, bcc: List[str] = None, 
                       attachments: List[Dict] = None) -> str:
        """Create MIME message for Yahoo Mail API"""
        message = MIMEMultipart('alternative')
        message['From'] = sender
        message['To'] = ', '.join(to)
        message['Subject'] = subject
        
        if bcc:
            message['Bcc'] = ', '.join(bcc)
        
        # Add HTML body
        html_part = MIMEText(body_html, 'html')
        message.attach(html_part)
        
        # Add attachments if provided
        if attachments:
            for attachment in attachments:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment['content'])
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {attachment["name"]}'
                )
                message.attach(part)
        
        # Encode message
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        return raw_message
    
    async def send_email(self, access_token: str, sender_email: str, 
                        to_emails: list, subject: str, body_html: str,
                        bcc_emails: list = None, attachments: list = None) -> Dict[str, Any]:
        """Send email via Yahoo Mail API"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Create message
        raw_message = self._create_message(
            sender_email, to_emails, subject, body_html, bcc_emails, attachments
        )
        
        # Prepare payload
        payload = {
            'raw': raw_message
        }
        
        # Yahoo Mail API endpoint (this may vary based on actual Yahoo API)
        url = f"{self.provider_config.api_base_url}/mail/v1/messages/send"
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload) as response:
                if response.status in [200, 201]:
                    response_data = await response.json()
                    return {
                        'success': True,
                        'status': response.status,
                        'message': 'Email sent successfully',
                        'message_id': response_data.get('id')
                    }
                elif response.status == 429:
                    # Rate limited
                    retry_after = int(response.headers.get('Retry-After', '60'))
                    raise OAuth2Error(f"Rate limited. Retry after {retry_after} seconds")
                else:
                    error_text = await response.text()
                    raise OAuth2Error(f"Send failed: {response.status} - {error_text}")
    
    async def get_profile(self, access_token: str) -> Dict[str, Any]:
        """Get Yahoo profile information"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/openid/v1/userinfo"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    return await response.json()
                else:
                    return {}
    
    async def test_connection(self, access_token: str) -> bool:
        """Test if the connection to Yahoo is working"""
        try:
            profile = await self.get_profile(access_token)
            return profile.get('email') is not None
        except OAuth2Error:
            return False
    
    async def get_mailbox_info(self, access_token: str) -> Dict[str, Any]:
        """Get Yahoo mailbox information"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # This endpoint may vary based on actual Yahoo API
        url = f"{self.provider_config.api_base_url}/mail/v1/mailbox"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    return await response.json()
                else:
                    return {}
    
    async def get_folders(self, access_token: str) -> List[Dict[str, Any]]:
        """Get Yahoo mail folders"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/mail/v1/folders"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    data = await response.json()
                    return data.get('folders', [])
                else:
                    return []
    
    async def get_messages_count(self, access_token: str, folder_id: str = 'inbox') -> int:
        """Get count of messages in folder"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/mail/v1/folders/{folder_id}/messages"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    data = await response.json()
                    return data.get('total', 0)
                else:
                    return 0
    
    async def create_draft(self, access_token: str, sender_email: str, 
                          to_emails: list, subject: str, body_html: str,
                          bcc_emails: list = None, attachments: list = None) -> Dict[str, Any]:
        """Create email draft"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Create message
        raw_message = self._create_message(
            sender_email, to_emails, subject, body_html, bcc_emails, attachments
        )
        
        # Prepare payload
        payload = {
            'message': {
                'raw': raw_message
            }
        }
        
        url = f"{self.provider_config.api_base_url}/mail/v1/drafts"
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload) as response:
                if response.status in [200, 201]:
                    response_data = await response.json()
                    return {
                        'success': True,
                        'draft_id': response_data.get('id'),
                        'message': 'Draft created successfully'
                    }
                else:
                    error_text = await response.text()
                    raise OAuth2Error(f"Draft creation failed: {response.status} - {error_text}")
    
    async def send_draft(self, access_token: str, draft_id: str) -> Dict[str, Any]:
        """Send existing draft"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/mail/v1/drafts/{draft_id}/send"
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers) as response:
                if response.status in [200, 201]:
                    response_data = await response.json()
                    return {
                        'success': True,
                        'message_id': response_data.get('id'),
                        'message': 'Draft sent successfully'
                    }
                else:
                    error_text = await response.text()
                    raise OAuth2Error(f"Draft send failed: {response.status} - {error_text}")
