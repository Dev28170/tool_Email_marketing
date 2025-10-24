"""
Database models and operations for Email Mass Sender
Handles account storage, email campaigns, and audit logging
"""

from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime, Boolean, JSON, Float, UniqueConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
from sqlalchemy.exc import SQLAlchemyError
import json

Base = declarative_base()

# Global database session
SessionLocal = None
engine = None

def init_db(database_url: str = None):
    """Initialize database connection"""
    global SessionLocal, engine
    
    if database_url is None:
        from config import config
        database_url = config.DATABASE_URL
    
    engine = create_engine(database_url, echo=False)
    SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
    
    # Create all tables
    Base.metadata.create_all(bind=engine)
    
    return SessionLocal

def get_db():
    """Get database session"""
    if SessionLocal is None:
        init_db()
    
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

class EmailAccount(Base):
    """Email account model"""
    __tablename__ = 'email_accounts'
    
    id = Column(Integer, primary_key=True)
    email = Column(String(255), unique=True, nullable=False)
    provider = Column(String(50), nullable=False)  # office365, gmail, yahoo, hotmail
    access_token = Column(Text, nullable=True)
    refresh_token = Column(Text, nullable=True)
    token_expires_at = Column(DateTime, nullable=True)
    is_active = Column(Boolean, default=True)
    is_verified = Column(Boolean, default=False)
    last_used = Column(DateTime, nullable=True)
    success_count = Column(Integer, default=0)
    error_count = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'id': self.id,
            'email': self.email,
            'provider': self.provider,
            'is_active': self.is_active,
            'is_verified': self.is_verified,
            'last_used': self.last_used.isoformat() if self.last_used else None,
            'success_count': self.success_count,
            'error_count': self.error_count,
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat()
        }
    
    def is_token_valid(self) -> bool:
        """Check if access token is still valid"""
        if not self.token_expires_at:
            return False
        return datetime.utcnow() < self.token_expires_at - timedelta(minutes=5)

class EmailCampaign(Base):
    """Email campaign model"""
    __tablename__ = 'email_campaigns'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    subject = Column(String(500), nullable=False)
    body_html = Column(Text, nullable=False)
    body_text = Column(Text, nullable=True)
    template_data = Column(JSON, nullable=True)  # Placeholder data
    status = Column(String(50), default='draft')  # draft, running, completed, paused, failed
    total_recipients = Column(Integer, default=0)
    sent_count = Column(Integer, default=0)
    failed_count = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)
    started_at = Column(DateTime, nullable=True)
    completed_at = Column(DateTime, nullable=True)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'id': self.id,
            'name': self.name,
            'subject': self.subject,
            'body_html': self.body_html,
            'body_text': self.body_text,
            'template_data': self.template_data,
            'status': self.status,
            'total_recipients': self.total_recipients,
            'sent_count': self.sent_count,
            'failed_count': self.failed_count,
            'created_at': self.created_at.isoformat(),
            'started_at': self.started_at.isoformat() if self.started_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None
        }

class EmailRecipient(Base):
    """Email recipient model"""
    __tablename__ = 'email_recipients'
    
    id = Column(Integer, primary_key=True)
    campaign_id = Column(Integer, nullable=False)
    email = Column(String(255), nullable=False)
    name = Column(String(255), nullable=True)
    custom_data = Column(JSON, nullable=True)  # Custom placeholder data
    status = Column(String(50), default='pending')  # pending, sent, failed, bounced
    sent_at = Column(DateTime, nullable=True)
    error_message = Column(Text, nullable=True)
    account_used = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'id': self.id,
            'campaign_id': self.campaign_id,
            'email': self.email,
            'name': self.name,
            'custom_data': self.custom_data,
            'status': self.status,
            'sent_at': self.sent_at.isoformat() if self.sent_at else None,
            'error_message': self.error_message,
            'account_used': self.account_used,
            'created_at': self.created_at.isoformat() if self.created_at else None
        }

class EmailLog(Base):
    """Email sending log model"""
    __tablename__ = 'email_logs'
    
    id = Column(Integer, primary_key=True)
    campaign_id = Column(Integer, nullable=True)
    recipient_id = Column(Integer, nullable=True)
    account_email = Column(String(255), nullable=False)
    recipient_email = Column(String(255), nullable=False)
    subject = Column(String(500), nullable=False)
    status = Column(String(50), nullable=False)  # success, failed, throttled
    error_message = Column(Text, nullable=True)
    response_data = Column(JSON, nullable=True)
    sent_at = Column(DateTime, default=datetime.utcnow)
    processing_time = Column(Float, nullable=True)  # Time taken to send
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'id': self.id,
            'campaign_id': self.campaign_id,
            'recipient_id': self.recipient_id,
            'account_email': self.account_email,
            'recipient_email': self.recipient_email,
            'subject': self.subject,
            'status': self.status,
            'error_message': self.error_message,
            'response_data': self.response_data,
            'sent_at': self.sent_at.isoformat(),
            'processing_time': self.processing_time
        }

class AppSetting(Base):
    """Simple key/value application settings store.
    Used to persist defaults like Direct Send subject/body across sessions.
    """
    __tablename__ = 'app_settings'
    __table_args__ = (
        UniqueConstraint('key', name='uq_app_settings_key'),
    )

    id = Column(Integer, primary_key=True)
    key = Column(String(255), nullable=False, unique=True)
    value = Column(Text, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'key': self.key,
            'value': self.value,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }

class SendingProgress(Base):
    """Real-time sending progress tracking model"""
    __tablename__ = 'sending_progress'
    
    id = Column(Integer, primary_key=True)
    campaign_id = Column(Integer, nullable=False)
    session_id = Column(String(255), nullable=False)  # Unique session identifier
    total_emails = Column(Integer, nullable=False, default=0)
    sent_count = Column(Integer, nullable=False, default=0)
    failed_count = Column(Integer, nullable=False, default=0)
    pending_count = Column(Integer, nullable=False, default=0)
    current_batch = Column(Integer, nullable=False, default=0)
    total_batches = Column(Integer, nullable=False, default=0)
    status = Column(String(50), nullable=False, default='pending')  # pending, running, completed, failed, cancelled
    start_time = Column(DateTime, default=datetime.utcnow)
    last_update = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    estimated_completion = Column(DateTime, nullable=True)
    current_speed = Column(Float, nullable=True)  # emails per minute
    error_message = Column(Text, nullable=True)
    progress_percentage = Column(Float, nullable=False, default=0.0)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'id': self.id,
            'campaign_id': self.campaign_id,
            'session_id': self.session_id,
            'total_emails': self.total_emails,
            'sent_count': self.sent_count,
            'failed_count': self.failed_count,
            'pending_count': self.pending_count,
            'current_batch': self.current_batch,
            'total_batches': self.total_batches,
            'status': self.status,
            'start_time': self.start_time.isoformat(),
            'last_update': self.last_update.isoformat(),
            'estimated_completion': self.estimated_completion.isoformat() if self.estimated_completion else None,
            'current_speed': self.current_speed,
            'error_message': self.error_message,
            'progress_percentage': self.progress_percentage
        }
    
    def calculate_progress(self):
        """Calculate progress percentage"""
        if self.total_emails > 0:
            completed = self.sent_count + self.failed_count
            self.progress_percentage = (completed / self.total_emails) * 100
        else:
            self.progress_percentage = 0.0
    
    def calculate_speed(self):
        """Calculate current sending speed (emails per minute)"""
        if self.start_time:
            elapsed_minutes = (datetime.utcnow() - self.start_time).total_seconds() / 60
            if elapsed_minutes > 0:
                completed = self.sent_count + self.failed_count
                self.current_speed = completed / elapsed_minutes
            else:
                self.current_speed = 0.0
    
    def estimate_completion(self):
        """Estimate completion time"""
        if self.current_speed and self.current_speed > 0:
            remaining_emails = self.total_emails - (self.sent_count + self.failed_count)
            if remaining_emails > 0:
                remaining_minutes = remaining_emails / self.current_speed
                self.estimated_completion = datetime.utcnow() + timedelta(minutes=remaining_minutes)
            else:
                self.estimated_completion = datetime.utcnow()

class DatabaseManager:
    """Database management class"""
    
    def __init__(self, database_url: str):
        self.engine = create_engine(database_url)
        self.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=self.engine)
        Base.metadata.create_all(bind=self.engine)
    
    def get_session(self) -> Session:
        """Get database session"""
        return self.SessionLocal()

    # ------------------------
    # App settings persistence
    # ------------------------
    def get_setting(self, key: str, default: str = None) -> Optional[str]:
        """Get a setting value by key."""
        with self.get_session() as session:
            row = session.query(AppSetting).filter(AppSetting.key == key).first()
            return row.value if row else default

    def set_setting(self, key: str, value: str) -> None:
        """Create or update a setting value by key."""
        with self.get_session() as session:
            row = session.query(AppSetting).filter(AppSetting.key == key).first()
            if row:
                row.value = value
                row.updated_at = datetime.utcnow()
            else:
                row = AppSetting(key=key, value=value)
                session.add(row)
            session.commit()

    def get_direct_send_defaults(self) -> Dict[str, Any]:
        """Convenience method to load Direct Send default subject/body."""
        return {
            'subject': self.get_setting('direct_send.subject', ''),
            'body': self.get_setting('direct_send.body', '')
        }

    def set_direct_send_defaults(self, subject: str, body: str) -> None:
        """Save Direct Send default subject/body."""
        self.set_setting('direct_send.subject', subject or '')
        self.set_setting('direct_send.body', body or '')

    # ------------------------
    # Campaign defaults (New Campaign form)
    # ------------------------
    def get_campaign_defaults(self) -> Dict[str, Any]:
        """Load default campaign subject/body used to prefill New Campaign form."""
        return {
            'subject': self.get_setting('campaign.default_subject', ''),
            'body': self.get_setting('campaign.default_body', '')
        }

    def set_campaign_defaults(self, subject: str, body: str) -> None:
        """Persist default campaign subject/body from last created campaign."""
        self.set_setting('campaign.default_subject', subject or '')
        self.set_setting('campaign.default_body', body or '')
    
    def add_email_account(self, email: str, provider: str, access_token: str = None, 
                         refresh_token: str = None, token_expires_at: datetime = None) -> EmailAccount:
        """Add new email account"""
        with self.get_session() as session:
            account = EmailAccount(
                email=email,
                provider=provider,
                access_token=access_token,
                refresh_token=refresh_token,
                token_expires_at=token_expires_at
            )
            session.add(account)
            session.commit()
            session.refresh(account)
            return account
    
    def get_active_accounts(self, provider: str = None) -> List[EmailAccount]:
        """Get active email accounts"""
        with self.get_session() as session:
            query = session.query(EmailAccount).filter(EmailAccount.is_active == True)
            if provider:
                query = query.filter(EmailAccount.provider == provider)
            return query.all()
    
    def update_account_tokens(self, account_id: int, access_token: str, 
                            refresh_token: str = None, expires_at: datetime = None):
        """Update account tokens"""
        with self.get_session() as session:
            account = session.query(EmailAccount).filter(EmailAccount.id == account_id).first()
            if account:
                account.access_token = access_token
                if refresh_token:
                    account.refresh_token = refresh_token
                if expires_at:
                    account.token_expires_at = expires_at
                account.updated_at = datetime.utcnow()
                session.commit()
    
    def increment_account_stats(self, account_id: int, success: bool = True):
        """Increment account success/error count"""
        with self.get_session() as session:
            account = session.query(EmailAccount).filter(EmailAccount.id == account_id).first()
            if account:
                if success:
                    account.success_count += 1
                else:
                    account.error_count += 1
                account.last_used = datetime.utcnow()
                session.commit()
    
    def create_campaign(self, name: str, subject: str, body_html: str, 
                       body_text: str = None, template_data: Dict = None) -> EmailCampaign:
        """Create new email campaign"""
        with self.get_session() as session:
            campaign = EmailCampaign(
                name=name,
                subject=subject,
                body_html=body_html,
                body_text=body_text,
                template_data=template_data
            )
            session.add(campaign)
            session.commit()
            session.refresh(campaign)
            return campaign
    
    def add_recipients(self, campaign_id: int, recipients: List[Dict[str, Any]]):
        """Add recipients to campaign"""
        with self.get_session() as session:
            for recipient_data in recipients:
                recipient = EmailRecipient(
                    campaign_id=campaign_id,
                    email=recipient_data['email'],
                    name=recipient_data.get('name'),
                    custom_data=recipient_data.get('custom_data')
                )
                session.add(recipient)
            session.commit()
    
    def log_email_send(self, account_email: str, recipient_email: str, subject: str,
                      status: str, error_message: str = None, response_data: Dict = None,
                      processing_time: float = None, campaign_id: int = None, 
                      recipient_id: int = None):
        """Log email sending attempt"""
        with self.get_session() as session:
            log = EmailLog(
                campaign_id=campaign_id,
                recipient_id=recipient_id,
                account_email=account_email,
                recipient_email=recipient_email,
                subject=subject,
                status=status,
                error_message=error_message,
                response_data=response_data,
                processing_time=processing_time
            )
            session.add(log)
            session.commit()
    
    def get_campaign_stats(self, campaign_id: int) -> Dict[str, Any]:
        """Get campaign statistics"""
        with self.get_session() as session:
            campaign = session.query(EmailCampaign).filter(EmailCampaign.id == campaign_id).first()
            if not campaign:
                return {}
            
            recipients = session.query(EmailRecipient).filter(EmailRecipient.campaign_id == campaign_id).all()
            
            stats = {
                'total_recipients': len(recipients),
                'sent': len([r for r in recipients if r.status == 'sent']),
                'failed': len([r for r in recipients if r.status == 'failed']),
                'pending': len([r for r in recipients if r.status == 'pending']),
                'bounced': len([r for r in recipients if r.status == 'bounced'])
            }
            
            return stats
    
    def get_all_campaigns(self) -> List[EmailCampaign]:
        """Get all campaigns"""
        with self.get_session() as session:
            return session.query(EmailCampaign).order_by(EmailCampaign.created_at.desc()).all()
    
    def create_sending_progress(self, campaign_id: int, session_id: str, total_emails: int, total_batches: int = 1) -> SendingProgress:
        """Create new sending progress record"""
        with self.get_session() as session:
            progress = SendingProgress(
                campaign_id=campaign_id,
                session_id=session_id,
                total_emails=total_emails,
                pending_count=total_emails,
                total_batches=total_batches,
                status='pending'
            )
            session.add(progress)
            session.commit()
            session.refresh(progress)
            return progress
    
    def update_sending_progress(self, session_id: str, **updates) -> SendingProgress:
        """Update sending progress"""
        with self.get_session() as session:
            progress = session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
            if progress:
                for key, value in updates.items():
                    if hasattr(progress, key):
                        setattr(progress, key, value)
                
                # Recalculate derived fields
                progress.calculate_progress()
                progress.calculate_speed()
                progress.estimate_completion()
                
                session.commit()
                session.refresh(progress)
            return progress
    
    def get_sending_progress(self, session_id: str) -> SendingProgress:
        """Get sending progress by session ID"""
        with self.get_session() as session:
            return session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
    
    def get_campaign_progress(self, campaign_id: int) -> SendingProgress:
        """Get latest sending progress for campaign"""
        with self.get_session() as session:
            return session.query(SendingProgress).filter(
                SendingProgress.campaign_id == campaign_id
            ).order_by(SendingProgress.last_update.desc()).first()
    
    def increment_sent_count(self, session_id: str) -> SendingProgress:
        """Increment sent count and update progress"""
        with self.get_session() as session:
            progress = session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
            if progress:
                progress.sent_count += 1
                progress.pending_count = max(0, progress.pending_count - 1)
                progress.calculate_progress()
                progress.calculate_speed()
                progress.estimate_completion()
                session.commit()
                session.refresh(progress)
            return progress
    
    def increment_failed_count(self, session_id: str) -> SendingProgress:
        """Increment failed count and update progress"""
        with self.get_session() as session:
            progress = session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
            if progress:
                progress.failed_count += 1
                progress.pending_count = max(0, progress.pending_count - 1)
                progress.calculate_progress()
                progress.calculate_speed()
                progress.estimate_completion()
                session.commit()
                session.refresh(progress)
            return progress
    
    def update_batch_progress(self, session_id: str, current_batch: int) -> SendingProgress:
        """Update current batch progress"""
        with self.get_session() as session:
            progress = session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
            if progress:
                progress.current_batch = current_batch
                session.commit()
                session.refresh(progress)
            return progress
    
    def complete_sending_progress(self, session_id: str, status: str = 'completed', error_message: str = None) -> SendingProgress:
        """Mark sending progress as completed"""
        with self.get_session() as session:
            progress = session.query(SendingProgress).filter(SendingProgress.session_id == session_id).first()
            if progress:
                progress.status = status
                progress.error_message = error_message
                progress.calculate_progress()
                progress.calculate_speed()
                progress.estimate_completion()
                session.commit()
                session.refresh(progress)
            return progress

# Global database manager instance
db_manager = None

def init_database(database_url: str):
    """Initialize database manager"""
    global db_manager
    db_manager = DatabaseManager(database_url)
    return db_manager

def get_db() -> DatabaseManager:
    """Get database manager instance"""
    if db_manager is None:
        raise RuntimeError("Database not initialized. Call init_database() first.")
    return db_manager
