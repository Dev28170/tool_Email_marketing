"""
Base authentication class for email providers
Provides common OAuth2.0 functionality and interface
"""

import asyncio
import aiohttp
import json
from abc import ABC, abstractmethod
from typing import Dict, Any, Optional, Tuple
from datetime import datetime, timedelta
from urllib.parse import urlencode, parse_qs, urlparse
import secrets
import hashlib
import base64
from config import config
from database import EmailAccount

class OAuth2Error(Exception):
    """OAuth2 authentication error"""
    pass

class TokenExpiredError(OAuth2Error):
    """Token has expired"""
    pass

class BaseOAuth2Provider(ABC):
    """Base class for OAuth2 email providers"""
    
    def __init__(self, provider_name: str):
        self.provider_name = provider_name
        self.provider_config = config.get_provider_config(provider_name)
        if not self.provider_config:
            raise ValueError(f"Configuration not found for provider: {provider_name}")
        
        # State management for CSRF protection
        self._state_storage = {}
    
    def generate_state(self) -> str:
        """Generate random state for CSRF protection"""
        state = secrets.token_urlsafe(32)
        self._state_storage[state] = datetime.utcnow()
        return state
    
    def validate_state(self, state: str) -> bool:
        """Validate state parameter"""
        if state not in self._state_storage:
            return False
        
        # Check if state is not too old (5 minutes)
        if datetime.utcnow() - self._state_storage[state] > timedelta(minutes=5):
            del self._state_storage[state]
            return False
        
        del self._state_storage[state]
        return True
    
    def get_auth_url(self, state: str = None) -> str:
        """Get OAuth2 authorization URL"""
        if not state:
            state = self.generate_state()
        
        params = {
            'client_id': self.provider_config.client_id,
            'redirect_uri': self.provider_config.redirect_uri,
            'scope': ' '.join(self.provider_config.scopes),
            'response_type': 'code',
            'state': state,
            'access_type': 'offline',
            'prompt': 'consent'
        }
        
        # Add provider-specific parameters
        params.update(self._get_auth_params())
        
        return f"{self.provider_config.auth_url}?{urlencode(params)}"
    
    @abstractmethod
    def _get_auth_params(self) -> Dict[str, str]:
        """Get provider-specific authorization parameters"""
        pass
    
    async def exchange_code_for_token(self, code: str, state: str = None) -> Dict[str, Any]:
        """Exchange authorization code for access token"""
        if state and not self.validate_state(state):
            raise OAuth2Error("Invalid or expired state parameter")
        
        token_data = {
            'client_id': self.provider_config.client_id,
            'client_secret': self.provider_config.client_secret,
            'code': code,
            'redirect_uri': self.provider_config.redirect_uri,
            'grant_type': 'authorization_code'
        }
        
        # Add provider-specific token parameters
        token_data.update(self._get_token_params())
        
        async with aiohttp.ClientSession() as session:
            async with session.post(
                self.provider_config.token_url,
                data=token_data,
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            ) as response:
                if response.status != 200:
                    error_text = await response.text()
                    raise OAuth2Error(f"Token exchange failed: {response.status} - {error_text}")
                
                token_response = await response.json()
                return self._process_token_response(token_response)
    
    @abstractmethod
    def _get_token_params(self) -> Dict[str, str]:
        """Get provider-specific token request parameters"""
        pass
    
    def _process_token_response(self, response: Dict[str, Any]) -> Dict[str, Any]:
        """Process token response and extract tokens"""
        if 'error' in response:
            raise OAuth2Error(f"Token error: {response['error']} - {response.get('error_description', '')}")
        
        access_token = response.get('access_token')
        refresh_token = response.get('refresh_token')
        expires_in = response.get('expires_in', 3600)
        
        if not access_token:
            raise OAuth2Error("No access token in response")
        
        expires_at = datetime.utcnow() + timedelta(seconds=expires_in)
        
        return {
            'access_token': access_token,
            'refresh_token': refresh_token,
            'expires_at': expires_at,
            'token_type': response.get('token_type', 'Bearer'),
            'scope': response.get('scope', '')
        }
    
    async def refresh_access_token(self, refresh_token: str) -> Dict[str, Any]:
        """Refresh access token using refresh token"""
        token_data = {
            'client_id': self.provider_config.client_id,
            'client_secret': self.provider_config.client_secret,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token'
        }
        
        # Add provider-specific refresh parameters
        token_data.update(self._get_refresh_params())
        
        async with aiohttp.ClientSession() as session:
            async with session.post(
                self.provider_config.token_url,
                data=token_data,
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            ) as response:
                if response.status != 200:
                    error_text = await response.text()
                    raise OAuth2Error(f"Token refresh failed: {response.status} - {error_text}")
                
                token_response = await response.json()
                return self._process_token_response(token_response)
    
    @abstractmethod
    def _get_refresh_params(self) -> Dict[str, str]:
        """Get provider-specific refresh token parameters"""
        pass
    
    async def get_user_info(self, access_token: str) -> Dict[str, Any]:
        """Get user information using access token"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        user_info_url = self._get_user_info_url()
        
        async with aiohttp.ClientSession() as session:
            async with session.get(user_info_url, headers=headers) as response:
                if response.status != 200:
                    error_text = await response.text()
                    raise OAuth2Error(f"Failed to get user info: {response.status} - {error_text}")
                
                return await response.json()
    
    @abstractmethod
    def _get_user_info_url(self) -> str:
        """Get user info endpoint URL"""
        pass
    
    async def validate_token(self, access_token: str) -> bool:
        """Validate if access token is still valid"""
        try:
            await self.get_user_info(access_token)
            return True
        except OAuth2Error:
            return False
    
    def is_token_expired(self, expires_at: datetime) -> bool:
        """Check if token is expired"""
        return datetime.utcnow() >= expires_at - timedelta(minutes=5)
    
    async def revoke_token(self, token: str) -> bool:
        """Revoke access token"""
        revoke_url = self._get_revoke_url()
        if not revoke_url:
            return False
        
        data = {'token': token}
        
        async with aiohttp.ClientSession() as session:
            async with session.post(revoke_url, data=data) as response:
                return response.status in [200, 204]
    
    @abstractmethod
    def _get_revoke_url(self) -> Optional[str]:
        """Get token revocation endpoint URL"""
        pass

class TokenManager:
    """Manages tokens for multiple accounts"""
    
    def __init__(self, db_manager):
        self.db_manager = db_manager
    
    async def get_valid_token(self, account_id: int, provider: str) -> str:
        """Get valid access token for account"""
        account = self.db_manager.get_session().query(EmailAccount).filter(
            EmailAccount.id == account_id
        ).first()
        
        if not account:
            raise OAuth2Error("Account not found")
        
        # Check if current token is valid
        if account.is_token_valid():
            return account.access_token
        
        # Try to refresh token
        if account.refresh_token:
            try:
                provider_auth = self._get_provider_auth(provider)
                new_tokens = await provider_auth.refresh_access_token(account.refresh_token)
                
                # Update tokens in database
                self.db_manager.update_account_tokens(
                    account_id,
                    new_tokens['access_token'],
                    new_tokens.get('refresh_token'),
                    new_tokens['expires_at']
                )
                
                return new_tokens['access_token']
            except OAuth2Error:
                pass
        
        # Token refresh failed, need re-authentication
        raise TokenExpiredError("Token expired and refresh failed")
    
    def _get_provider_auth(self, provider: str) -> BaseOAuth2Provider:
        """Get provider authentication instance"""
        from .office365_auth import Office365Auth
        from .gmail_auth import GmailAuth
        from .yahoo_auth import YahooAuth
        from .hotmail_auth import HotmailAuth
        
        providers = {
            'office365': Office365Auth,
            'gmail': GmailAuth,
            'yahoo': YahooAuth,
            'hotmail': HotmailAuth
        }
        
        if provider not in providers:
            raise ValueError(f"Unsupported provider: {provider}")
        
        return providers[provider]()
