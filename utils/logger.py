"""
Advanced logging system for Email Mass Sender
Provides structured logging with different levels and outputs
"""

import os
import sys
from datetime import datetime
from typing import Dict, Any, Optional
from loguru import logger
from config import config

class EmailLogger:
    """Advanced logging system for email operations"""
    
    def __init__(self):
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging configuration"""
        # Remove default handler
        logger.remove()
        
        # Console logging
        logger.add(
            sys.stdout,
            level=config.LOG_LEVEL,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
            colorize=True
        )
        
        # File logging
        os.makedirs(os.path.dirname(config.LOG_FILE), exist_ok=True)
        logger.add(
            config.LOG_FILE,
            level=config.LOG_LEVEL,
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            rotation="1 day",
            retention="30 days",
            compression="zip"
        )
        
        # Error logging
        error_log_file = config.LOG_FILE.replace('.log', '_errors.log')
        logger.add(
            error_log_file,
            level="ERROR",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            rotation="1 week",
            retention="12 weeks",
            compression="zip"
        )
        
        # Email operation logging
        email_log_file = config.LOG_FILE.replace('.log', '_emails.log')
        logger.add(
            email_log_file,
            level="INFO",
            format="{time:YYYY-MM-DD HH:mm:ss} | {message}",
            rotation="1 day",
            retention="7 days",
            compression="zip",
            filter=lambda record: "EMAIL_OP" in record["extra"]
        )
    
    def log_email_operation(self, operation: str, details: Dict[str, Any], level: str = "INFO"):
        """Log email operation with structured data"""
        log_data = {
            "operation": operation,
            "timestamp": datetime.utcnow().isoformat(),
            **details
        }
        
        logger.bind(EMAIL_OP=True).log(
            level,
            f"EMAIL_OP: {operation} | {self._format_log_data(log_data)}"
        )
    
    def log_authentication(self, provider: str, email: str, success: bool, details: Dict[str, Any] = None):
        """Log authentication events"""
        log_data = {
            "provider": provider,
            "email": email,
            "success": success,
            "timestamp": datetime.utcnow().isoformat(),
            **(details or {})
        }
        
        level = "INFO" if success else "ERROR"
        logger.log(
            level,
            f"AUTH: {provider} | {email} | {'SUCCESS' if success else 'FAILED'} | {self._format_log_data(log_data)}"
        )
    
    def log_campaign_event(self, campaign_id: int, event: str, details: Dict[str, Any] = None):
        """Log campaign events"""
        log_data = {
            "campaign_id": campaign_id,
            "event": event,
            "timestamp": datetime.utcnow().isoformat(),
            **(details or {})
        }
        
        logger.info(f"CAMPAIGN: {campaign_id} | {event} | {self._format_log_data(log_data)}")
    
    def log_rate_limit(self, provider: str, account: str, retry_after: int):
        """Log rate limiting events"""
        log_data = {
            "provider": provider,
            "account": account,
            "retry_after": retry_after,
            "timestamp": datetime.utcnow().isoformat()
        }
        
        logger.warning(f"RATE_LIMIT: {provider} | {account} | Retry after {retry_after}s | {self._format_log_data(log_data)}")
    
    def log_error(self, error: Exception, context: Dict[str, Any] = None):
        """Log errors with context"""
        log_data = {
            "error_type": type(error).__name__,
            "error_message": str(error),
            "timestamp": datetime.utcnow().isoformat(),
            **(context or {})
        }
        
        logger.error(f"ERROR: {type(error).__name__} | {str(error)} | {self._format_log_data(log_data)}")
    
    def log_performance(self, operation: str, duration: float, details: Dict[str, Any] = None):
        """Log performance metrics"""
        log_data = {
            "operation": operation,
            "duration": duration,
            "timestamp": datetime.utcnow().isoformat(),
            **(details or {})
        }
        
        logger.info(f"PERFORMANCE: {operation} | {duration:.3f}s | {self._format_log_data(log_data)}")
    
    def _format_log_data(self, data: Dict[str, Any]) -> str:
        """Format log data as string"""
        formatted_items = []
        for key, value in data.items():
            if isinstance(value, (dict, list)):
                formatted_items.append(f"{key}={str(value)}")
            else:
                formatted_items.append(f"{key}={value}")
        return " | ".join(formatted_items)
    
    def get_log_stats(self) -> Dict[str, Any]:
        """Get logging statistics"""
        try:
            log_file = config.LOG_FILE
            if not os.path.exists(log_file):
                return {"error": "Log file not found"}
            
            # Get file size
            file_size = os.path.getsize(log_file)
            
            # Count log entries by level (approximate)
            with open(log_file, 'r') as f:
                content = f.read()
                error_count = content.count('| ERROR |')
                warning_count = content.count('| WARNING |')
                info_count = content.count('| INFO |')
                debug_count = content.count('| DEBUG |')
            
            return {
                "log_file": log_file,
                "file_size_mb": round(file_size / (1024 * 1024), 2),
                "error_count": error_count,
                "warning_count": warning_count,
                "info_count": info_count,
                "debug_count": debug_count,
                "total_entries": error_count + warning_count + info_count + debug_count
            }
        except Exception as e:
            return {"error": str(e)}

# Global logger instance
email_logger = EmailLogger()

# Convenience functions
def log_email_send(account: str, recipient: str, success: bool, details: Dict[str, Any] = None):
    """Log email send operation"""
    email_logger.log_email_operation(
        "SEND_EMAIL",
        {
            "account": account,
            "recipient": recipient,
            "success": success,
            **(details or {})
        }
    )

def log_campaign_start(campaign_id: int, recipient_count: int):
    """Log campaign start"""
    email_logger.log_campaign_event(
        campaign_id,
        "STARTED",
        {"recipient_count": recipient_count}
    )

def log_campaign_complete(campaign_id: int, sent: int, failed: int, duration: float):
    """Log campaign completion"""
    email_logger.log_campaign_event(
        campaign_id,
        "COMPLETED",
        {
            "sent": sent,
            "failed": failed,
            "duration": duration
        }
    )

def log_auth_success(provider: str, email: str):
    """Log successful authentication"""
    email_logger.log_authentication(provider, email, True)

def log_auth_failure(provider: str, email: str, error: str):
    """Log failed authentication"""
    email_logger.log_authentication(provider, email, False, {"error": error})

def log_rate_limit_hit(provider: str, account: str, retry_after: int):
    """Log rate limit hit"""
    email_logger.log_rate_limit(provider, account, retry_after)

def log_performance_metric(operation: str, duration: float, details: Dict[str, Any] = None):
    """Log performance metric"""
    email_logger.log_performance(operation, duration, details)
