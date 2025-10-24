"""
Configuration management for Email Mass Sender
Handles environment variables, settings, and provider configurations
"""

import os
from typing import Dict, Any, Optional
from dataclasses import dataclass
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

@dataclass
class ProviderConfig:
    """Configuration for email providers"""
    client_id: str
    client_secret: str
    redirect_uri: str
    scopes: list
    auth_url: str
    token_url: str
    api_base_url: str

class Config:
    """Main configuration class"""
    
    # Application Settings
    SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key-change-this')
    DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
    HOST = os.getenv('HOST', '127.0.0.1')
    PORT = int(os.getenv('PORT', 8080))
    
    # Database
    DATABASE_URL = os.getenv('DATABASE_URL', 'sqlite:///email_sender.db')
    
    # OpenAI Configuration
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-4o-mini')
    OPENAI_MAX_TOKENS = int(os.getenv('OPENAI_MAX_TOKENS', 200))
    
    # Concurrency Settings - Optimized for Speed
    MAX_CONCURRENT_ACCOUNTS = int(os.getenv('MAX_CONCURRENT_ACCOUNTS', 500))  # Increased from 200
    MAX_CONCURRENT_PER_PROVIDER = int(os.getenv('MAX_CONCURRENT_PER_PROVIDER', 100))  # Increased from 50
    REQUEST_TIMEOUT = int(os.getenv('REQUEST_TIMEOUT', 15))  # Reduced from 30
    RETRY_ATTEMPTS = int(os.getenv('RETRY_ATTEMPTS', 3))  # Reduced from 5
    RETRY_DELAY = float(os.getenv('RETRY_DELAY', 0.5))  # Reduced from 1.0
    
    # Rate Limiting - Optimized for Speed
    RATE_LIMIT_PER_MINUTE = int(os.getenv('RATE_LIMIT_PER_MINUTE', 200))  # Increased from 100
    RATE_LIMIT_BURST = int(os.getenv('RATE_LIMIT_BURST', 20))  # Increased from 10
    
    # Speed Optimization Settings
    BATCH_SIZE_OPTIMIZED = int(os.getenv('BATCH_SIZE_OPTIMIZED', 100))  # Increased from 50
    MAX_CONCURRENT_BATCHES = int(os.getenv('MAX_CONCURRENT_BATCHES', 10))  # New setting
    UI_INTERACTION_TIMEOUT = int(os.getenv('UI_INTERACTION_TIMEOUT', 2000))  # New setting
    SEND_CONFIRMATION_TIMEOUT = int(os.getenv('SEND_CONFIRMATION_TIMEOUT', 3000))  # New setting
    NETWORK_IDLE_TIMEOUT = int(os.getenv('NETWORK_IDLE_TIMEOUT', 5000))  # New setting
    
    # File Upload
    MAX_ATTACHMENT_SIZE = int(os.getenv('MAX_ATTACHMENT_SIZE', 25 * 1024 * 1024))  # 25MB
    UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
    ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'doc', 'docx', 'xls', 'xlsx'}
    
    # Logging
    LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
    LOG_FILE = os.getenv('LOG_FILE', 'logs/email_sender.log')
    
    @classmethod
    def get_provider_config(cls, provider: str) -> Optional[ProviderConfig]:
        """Get configuration for specific email provider"""
        configs = {
            'office365': ProviderConfig(
                client_id=os.getenv('OFFICE365_CLIENT_ID', ''),
                client_secret=os.getenv('OFFICE365_CLIENT_SECRET', ''),
                redirect_uri=os.getenv('OFFICE365_REDIRECT_URI', 'http://localhost:8080/callback/office365'),
                scopes=[
                    'https://graph.microsoft.com/Mail.Send',
                    'https://graph.microsoft.com/Mail.ReadWrite',
                    'https://graph.microsoft.com/User.Read'
                ],
                auth_url='https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
                token_url='https://login.microsoftonline.com/common/oauth2/v2.0/token',
                api_base_url='https://graph.microsoft.com/v1.0'
            ),
            'gmail': ProviderConfig(
                client_id=os.getenv('GMAIL_CLIENT_ID', ''),
                client_secret=os.getenv('GMAIL_CLIENT_SECRET', ''),
                redirect_uri=os.getenv('GMAIL_REDIRECT_URI', 'http://localhost:8080/callback/gmail'),
                scopes=[
                    'https://www.googleapis.com/auth/gmail.send',
                    'https://www.googleapis.com/auth/gmail.compose',
                    'https://www.googleapis.com/auth/userinfo.email'
                ],
                auth_url='https://accounts.google.com/o/oauth2/v2/auth',
                token_url='https://oauth2.googleapis.com/token',
                api_base_url='https://gmail.googleapis.com/gmail/v1'
            ),
            'yahoo': ProviderConfig(
                client_id=os.getenv('YAHOO_CLIENT_ID', ''),
                client_secret=os.getenv('YAHOO_CLIENT_SECRET', ''),
                redirect_uri=os.getenv('YAHOO_REDIRECT_URI', 'http://localhost:8080/callback/yahoo'),
                scopes=['mail-w', 'mail-r'],
                auth_url='https://api.login.yahoo.com/oauth2/request_auth',
                token_url='https://api.login.yahoo.com/oauth2/get_token',
                api_base_url='https://api.login.yahoo.com'
            ),
            'hotmail': ProviderConfig(
                client_id=os.getenv('HOTMAIL_CLIENT_ID', ''),
                client_secret=os.getenv('HOTMAIL_CLIENT_SECRET', ''),
                redirect_uri=os.getenv('HOTMAIL_REDIRECT_URI', 'http://localhost:8080/callback/hotmail'),
                scopes=[
                    'https://graph.microsoft.com/Mail.Send',
                    'https://graph.microsoft.com/Mail.ReadWrite',
                    'https://graph.microsoft.com/User.Read'
                ],
                auth_url='https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
                token_url='https://login.microsoftonline.com/common/oauth2/v2.0/token',
                api_base_url='https://graph.microsoft.com/v1.0'
            )
        }
        return configs.get(provider.lower())
    
    @classmethod
    def validate_config(cls) -> Dict[str, Any]:
        """Validate configuration and return status"""
        validation = {
            'valid': True,
            'errors': [],
            'warnings': []
        }
        
        # Check required settings
        if not cls.SECRET_KEY or cls.SECRET_KEY == 'your-secret-key-change-this':
            validation['warnings'].append('SECRET_KEY should be changed for production')
        
        if not cls.OPENAI_API_KEY:
            validation['warnings'].append('OPENAI_API_KEY not set - AI features will be disabled')
        
        # Check provider configurations
        providers = ['office365', 'gmail', 'yahoo', 'hotmail']
        for provider in providers:
            config = cls.get_provider_config(provider)
            if not config or not config.client_id or not config.client_secret:
                validation['warnings'].append(f'{provider.upper()} credentials not configured')
        
        return validation

def init_database(database_url: str = None):
    """Initialize database manager and return it"""
    if database_url is None:
        database_url = Config.DATABASE_URL
    from database import init_database as _init_db_manager
    return _init_db_manager(database_url)

# Global config instance
config = Config()
