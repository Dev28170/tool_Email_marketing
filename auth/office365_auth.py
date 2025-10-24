"""
Office365 OAuth2.0 authentication implementation
Handles Microsoft Graph API authentication and token management
"""

import asyncio
import aiohttp
from typing import Dict, Any, Optional
from datetime import datetime, timedelta
from .base_auth import BaseOAuth2Provider, OAuth2Error
from config import config

class Office365Auth(BaseOAuth2Provider):
    """Office365 OAuth2.0 authentication provider"""
    
    def __init__(self):
        super().__init__('office365')
    
    def _get_auth_params(self) -> Dict[str, str]:
        """Get Office365-specific authorization parameters"""
        return {
            'response_mode': 'query',
            'login_hint': '',  # Can be pre-filled with email
        }
    
    def _get_token_params(self) -> Dict[str, str]:
        """Get Office365-specific token request parameters"""
        return {}
    
    def _get_refresh_params(self) -> Dict[str, str]:
        """Get Office365-specific refresh token parameters"""
        return {}
    
    def _get_user_info_url(self) -> str:
        """Get Office365 user info endpoint"""
        return f"{self.provider_config.api_base_url}/me"
    
    def _get_revoke_url(self) -> Optional[str]:
        """Get Office365 token revocation endpoint"""
        return "https://login.microsoftonline.com/common/oauth2/v2.0/logout"
    
    async def get_user_info(self, access_token: str) -> Dict[str, Any]:
        """Get Office365 user information"""
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
                    'id': user_data.get('id'),
                    'email': user_data.get('mail') or user_data.get('userPrincipalName'),
                    'name': user_data.get('displayName'),
                    'first_name': user_data.get('givenName'),
                    'last_name': user_data.get('surname'),
                    'provider': 'office365'
                }
    
    async def send_email(self, access_token: str, sender_email: str, 
                        to_emails: list, subject: str, body_html: str,
                        bcc_emails: list = None, attachments: list = None) -> Dict[str, Any]:
        """Send email via Microsoft Graph API"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Prepare recipients
        to_recipients = [{'emailAddress': {'address': email}} for email in to_emails]
        bcc_recipients = [{'emailAddress': {'address': email}} for email in (bcc_emails or [])]
        
        # Prepare message payload
        message = {
            'subject': subject,
            'body': {
                'contentType': 'HTML',
                'content': body_html
            },
            'toRecipients': to_recipients
        }
        
        if bcc_recipients:
            message['bccRecipients'] = bcc_recipients
        
        # Add attachments if provided
        if attachments:
            message['attachments'] = []
            for attachment in attachments:
                if attachment.get('size', 0) < 4 * 1024 * 1024:  # < 4MB
                    # Small attachment - include directly
                    import base64
                    content_bytes = base64.b64encode(attachment['content']).decode('utf-8')
                    message['attachments'].append({
                        '@odata.type': '#microsoft.graph.fileAttachment',
                        'name': attachment['name'],
                        'contentBytes': content_bytes
                    })
                else:
                    # Large attachment - use upload session
                    attachment_id = await self._upload_large_attachment(
                        access_token, sender_email, attachment
                    )
                    message['attachments'].append({
                        '@odata.type': '#microsoft.graph.referenceAttachment',
                        'id': attachment_id
                    })
        
        # Send email
        payload = {
            'message': message,
            'saveToSentItems': True
        }
        
        url = f"{self.provider_config.api_base_url}/users/{sender_email}/sendMail"
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload) as response:
                if response.status in [200, 202]:
                    return {
                        'success': True,
                        'status': response.status,
                        'message': 'Email sent successfully'
                    }
                elif response.status == 429:
                    # Rate limited
                    retry_after = int(response.headers.get('Retry-After', '60'))
                    raise OAuth2Error(f"Rate limited. Retry after {retry_after} seconds")
                else:
                    error_text = await response.text()
                    raise OAuth2Error(f"Send failed: {response.status} - {error_text}")
    
    async def _upload_large_attachment(self, access_token: str, sender_email: str, 
                                     attachment: Dict[str, Any]) -> str:
        """Upload large attachment using upload session"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Create upload session
        upload_session_payload = {
            'AttachmentItem': {
                'attachmentType': 'file',
                'name': attachment['name'],
                'size': len(attachment['content'])
            }
        }
        
        url = f"{self.provider_config.api_base_url}/users/{sender_email}/messages/createUploadSession"
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=upload_session_payload) as response:
                if response.status not in [200, 201]:
                    error_text = await response.text()
                    raise OAuth2Error(f"Failed to create upload session: {response.status} - {error_text}")
                
                session_data = await response.json()
                upload_url = session_data.get('uploadUrl')
                
                if not upload_url:
                    raise OAuth2Error("No upload URL in session response")
                
                # Upload file in chunks
                chunk_size = 4 * 1024 * 1024  # 4MB chunks
                content = attachment['content']
                total_size = len(content)
                
                for start in range(0, total_size, chunk_size):
                    end = min(start + chunk_size - 1, total_size - 1)
                    chunk = content[start:end + 1]
                    
                    chunk_headers = {
                        'Content-Length': str(len(chunk)),
                        'Content-Range': f'bytes {start}-{end}/{total_size}'
                    }
                    
                    async with session.put(upload_url, headers=chunk_headers, data=chunk) as chunk_response:
                        if chunk_response.status not in [200, 201, 202]:
                            error_text = await chunk_response.text()
                            raise OAuth2Error(f"Upload chunk failed: {chunk_response.status} - {error_text}")
                
                # Get final attachment info
                async with session.get(upload_url) as final_response:
                    if final_response.status == 200:
                        final_data = await final_response.json()
                        return final_data.get('id', '')
                    else:
                        raise OAuth2Error("Failed to get final attachment info")
    
    async def test_connection(self, access_token: str) -> bool:
        """Test if the connection to Office365 is working"""
        try:
            user_info = await self.get_user_info(access_token)
            return user_info.get('email') is not None
        except OAuth2Error:
            return False
    
    async def get_mailbox_info(self, access_token: str, email: str) -> Dict[str, Any]:
        """Get mailbox information"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/users/{email}/mailboxSettings"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    return await response.json()
                else:
                    return {}
    
    async def get_sent_items_count(self, access_token: str, email: str) -> int:
        """Get count of sent items"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{self.provider_config.api_base_url}/users/{email}/mailFolders/sentItems/messages/$count"
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status == 200:
                    return int(await response.text())
                else:
                    return 0
