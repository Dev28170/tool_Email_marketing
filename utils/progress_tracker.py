"""
Real-time Progress Tracking Service
Handles progress tracking for email sending operations
"""

import uuid
import time
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
from loguru import logger
import threading
from collections import defaultdict


class ProgressTracker:
    """Real-time progress tracking for email sending operations"""
    
    def __init__(self, db_manager):
        """Initialize progress tracker"""
        self.db_manager = db_manager
        self.active_sessions: Dict[str, Dict[str, Any]] = {}
        self.session_lock = threading.Lock()
        
    def create_session(self, campaign_id: int, total_emails: int, total_batches: int = 1) -> str:
        """Create a new progress tracking session"""
        session_id = str(uuid.uuid4())
        
        with self.session_lock:
            # Create database record
            progress = self.db_manager.create_sending_progress(
                campaign_id=campaign_id,
                session_id=session_id,
                total_emails=total_emails,
                total_batches=total_batches
            )
            
            # Store in memory for quick access
            self.active_sessions[session_id] = {
                'campaign_id': campaign_id,
                'total_emails': total_emails,
                'total_batches': total_batches,
                'sent_count': 0,
                'failed_count': 0,
                'pending_count': total_emails,
                'current_batch': 0,
                'status': 'pending',
                'start_time': datetime.utcnow(),
                'last_update': datetime.utcnow(),
                'progress_percentage': 0.0,
                'current_speed': 0.0,
                'estimated_completion': None,
                'error_message': None
            }
            
            logger.info(f"Created progress session {session_id} for campaign {campaign_id} with {total_emails} emails")
            return session_id
    
    def start_session(self, session_id: str) -> bool:
        """Mark session as started"""
        with self.session_lock:
            if session_id in self.active_sessions:
                self.active_sessions[session_id]['status'] = 'running'
                self.active_sessions[session_id]['start_time'] = datetime.utcnow()
                
                # Update database
                self.db_manager.update_sending_progress(session_id, status='running')
                
                logger.info(f"Started progress session {session_id}")
                return True
            return False
    
    def update_progress(self, session_id: str, sent_count: int = None, failed_count: int = None, 
                       current_batch: int = None, status: str = None, error_message: str = None) -> Dict[str, Any]:
        """Update progress for a session"""
        with self.session_lock:
            if session_id not in self.active_sessions:
                logger.warning(f"Session {session_id} not found")
                return {}
            
            session_data = self.active_sessions[session_id]
            
            # Update counts
            if sent_count is not None:
                session_data['sent_count'] = sent_count
            if failed_count is not None:
                session_data['failed_count'] = failed_count
            if current_batch is not None:
                session_data['current_batch'] = current_batch
            if status is not None:
                session_data['status'] = status
            if error_message is not None:
                session_data['error_message'] = error_message
            
            # Recalculate derived values
            session_data['pending_count'] = session_data['total_emails'] - session_data['sent_count'] - session_data['failed_count']
            session_data['last_update'] = datetime.utcnow()
            
            # Calculate progress percentage
            if session_data['total_emails'] > 0:
                completed = session_data['sent_count'] + session_data['failed_count']
                session_data['progress_percentage'] = (completed / session_data['total_emails']) * 100
            else:
                session_data['progress_percentage'] = 0.0
            
            # Calculate speed
            if session_data['start_time']:
                # Handle both datetime objects and ISO strings
                start_time = session_data['start_time']
                if isinstance(start_time, str):
                    start_time = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
                
                elapsed_minutes = (datetime.utcnow() - start_time).total_seconds() / 60
                if elapsed_minutes > 0:
                    completed = session_data['sent_count'] + session_data['failed_count']
                    session_data['current_speed'] = completed / elapsed_minutes
                else:
                    session_data['current_speed'] = 0.0
            
            # Estimate completion
            if session_data['current_speed'] and session_data['current_speed'] > 0:
                remaining_emails = session_data['total_emails'] - (session_data['sent_count'] + session_data['failed_count'])
                if remaining_emails > 0:
                    remaining_minutes = remaining_emails / session_data['current_speed']
                    session_data['estimated_completion'] = datetime.utcnow() + timedelta(minutes=remaining_minutes)
                else:
                    session_data['estimated_completion'] = datetime.utcnow()
            
            # Ensure all datetime objects are converted to ISO strings for JSON serialization
            for key, value in session_data.items():
                if isinstance(value, datetime):
                    session_data[key] = value.isoformat()
            
            # Update database
            self.db_manager.update_sending_progress(
                session_id,
                sent_count=session_data['sent_count'],
                failed_count=session_data['failed_count'],
                pending_count=session_data['pending_count'],
                current_batch=session_data['current_batch'],
                status=session_data['status'],
                error_message=session_data['error_message']
            )
            
            return session_data.copy()
    
    def increment_sent(self, session_id: str) -> Dict[str, Any]:
        """Increment sent count"""
        with self.session_lock:
            if session_id in self.active_sessions:
                self.active_sessions[session_id]['sent_count'] += 1
                return self.update_progress(session_id)
            return {}
    
    def increment_failed(self, session_id: str) -> Dict[str, Any]:
        """Increment failed count"""
        with self.session_lock:
            if session_id in self.active_sessions:
                self.active_sessions[session_id]['failed_count'] += 1
                return self.update_progress(session_id)
            return {}
    
    def update_batch(self, session_id: str, current_batch: int) -> Dict[str, Any]:
        """Update current batch"""
        return self.update_progress(session_id, current_batch=current_batch)
    
    def complete_session(self, session_id: str, status: str = 'completed', error_message: str = None) -> Dict[str, Any]:
        """Mark session as completed"""
        with self.session_lock:
            if session_id in self.active_sessions:
                try:
                    # Force update the session status immediately
                    session_data = self.active_sessions[session_id]
                    session_data['status'] = status
                    if error_message:
                        session_data['error_message'] = error_message
                    
                    # Calculate final progress
                    if status == 'completed':
                        session_data['progress_percentage'] = 100.0
                        session_data['estimated_completion'] = datetime.utcnow()
                    
                    # Update database
                    self.db_manager.complete_sending_progress(session_id, status, error_message)
                    
                    logger.info(f"Completed progress session {session_id} with status {status}")
                    
                    # Return the updated session data
                    return session_data.copy()
                except Exception as e:
                    logger.error(f"Error completing session {session_id}: {e}")
                    # Return basic completion data even if update fails
                    return {
                        'session_id': session_id,
                        'status': status,
                        'error_message': error_message,
                        'progress_percentage': 100.0 if status == 'completed' else 0.0
                    }
            else:
                logger.warning(f"Session {session_id} not found in active sessions for completion")
                return {}
    
    def cancel_session(self, session_id: str) -> bool:
        """Cancel a session"""
        with self.session_lock:
            if session_id in self.active_sessions:
                self.complete_session(session_id, status='cancelled')
                logger.info(f"Cancelled progress session {session_id}")
                return True
            return False
    
    def get_progress(self, session_id: str) -> Dict[str, Any]:
        """Get current progress for a session"""
        with self.session_lock:
            if session_id in self.active_sessions:
                session_data = self.active_sessions[session_id].copy()
                # Ensure all datetime objects are converted to ISO strings
                for key, value in session_data.items():
                    if isinstance(value, datetime):
                        session_data[key] = value.isoformat()
                return session_data
            else:
                # Try to get from database
                progress = self.db_manager.get_sending_progress(session_id)
                if progress:
                    return progress.to_dict()
                return {}
    
    def get_campaign_progress(self, campaign_id: int) -> Dict[str, Any]:
        """Get latest progress for a campaign"""
        progress = self.db_manager.get_campaign_progress(campaign_id)
        if progress:
            return progress.to_dict()
        return {}
    
    def cleanup_old_sessions(self, max_age_hours: int = 24):
        """Clean up old completed sessions from memory"""
        with self.session_lock:
            cutoff_time = datetime.utcnow() - timedelta(hours=max_age_hours)
            sessions_to_remove = []
            
            for session_id, session_data in self.active_sessions.items():
                if (session_data['status'] in ['completed', 'failed', 'cancelled'] and 
                    session_data['last_update'] < cutoff_time):
                    sessions_to_remove.append(session_id)
            
            for session_id in sessions_to_remove:
                del self.active_sessions[session_id]
                logger.info(f"Cleaned up old session {session_id}")
    
    def get_active_sessions(self) -> List[Dict[str, Any]]:
        """Get all active sessions"""
        with self.session_lock:
            return [session_data.copy() for session_data in self.active_sessions.values()]
    
    def get_session_stats(self) -> Dict[str, Any]:
        """Get overall statistics"""
        with self.session_lock:
            total_sessions = len(self.active_sessions)
            running_sessions = sum(1 for s in self.active_sessions.values() if s['status'] == 'running')
            completed_sessions = sum(1 for s in self.active_sessions.values() if s['status'] == 'completed')
            failed_sessions = sum(1 for s in self.active_sessions.values() if s['status'] == 'failed')
            
            return {
                'total_sessions': total_sessions,
                'running_sessions': running_sessions,
                'completed_sessions': completed_sessions,
                'failed_sessions': failed_sessions,
                'active_sessions': running_sessions
            }


# Global progress tracker instance
progress_tracker = None

def init_progress_tracker(db_manager):
    """Initialize global progress tracker"""
    global progress_tracker
    progress_tracker = ProgressTracker(db_manager)
    return progress_tracker

def get_progress_tracker():
    """Get global progress tracker instance"""
    if progress_tracker is None:
        raise RuntimeError("Progress tracker not initialized. Call init_progress_tracker() first.")
    return progress_tracker
