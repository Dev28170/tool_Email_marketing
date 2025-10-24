"""
Attachment handling system for email mass sender
Supports file upload, validation, and processing for all email providers
"""

import os
import mimetypes
import hashlib
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from pathlib import Path
import aiofiles
from loguru import logger
from config import config

class AttachmentError(Exception):
    """Attachment processing error"""
    pass

class AttachmentValidator:
    """Validates file attachments"""
    
    @staticmethod
    def validate_file_size(file_size: int) -> bool:
        """Validate file size"""
        return file_size <= config.MAX_ATTACHMENT_SIZE
    
    @staticmethod
    def validate_file_extension(filename: str) -> bool:
        """Validate file extension"""
        if not filename:
            return False
        
        extension = Path(filename).suffix.lower().lstrip('.')
        return extension in config.ALLOWED_EXTENSIONS
    
    @staticmethod
    def validate_mime_type(file_path: str) -> bool:
        """Validate MIME type"""
        mime_type, _ = mimetypes.guess_type(file_path)
        if not mime_type:
            return False
        
        # Common allowed MIME types
        allowed_mime_types = {
            'text/plain', 'text/html', 'text/csv',
            'application/pdf', 'application/msword', 
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/zip', 'application/x-zip-compressed',
            'image/jpeg', 'image/jpg', 'image/png', 'image/gif', 'image/webp',
            'audio/mpeg', 'audio/wav', 'audio/ogg',
            'video/mp4', 'video/avi', 'video/mov'
        }
        
        return mime_type in allowed_mime_types
    
    @staticmethod
    def validate_filename(filename: str) -> bool:
        """Validate filename"""
        if not filename:
            return False
        
        # Check for dangerous characters
        dangerous_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
        if any(char in filename for char in dangerous_chars):
            return False
        
        # Check filename length
        if len(filename) > 255:
            return False
        
        return True
    
    @classmethod
    def validate_attachment(cls, file_path: str, filename: str) -> Tuple[bool, str]:
        """Comprehensive attachment validation"""
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                return False, "File does not exist"
            
            # Check file size
            file_size = os.path.getsize(file_path)
            if not cls.validate_file_size(file_size):
                max_size_mb = config.MAX_ATTACHMENT_SIZE / (1024 * 1024)
                return False, f"File size exceeds maximum allowed size of {max_size_mb:.1f}MB"
            
            # Check filename
            if not cls.validate_filename(filename):
                return False, "Invalid filename"
            
            # Check file extension
            if not cls.validate_file_extension(filename):
                allowed_extensions = ', '.join(config.ALLOWED_EXTENSIONS)
                return False, f"File type not allowed. Allowed types: {allowed_extensions}"
            
            # Check MIME type
            if not cls.validate_mime_type(file_path):
                return False, "File type not allowed based on MIME type"
            
            return True, "Valid"
        
        except Exception as e:
            return False, f"Validation error: {str(e)}"

class AttachmentProcessor:
    """Processes and manages file attachments"""
    
    def __init__(self, upload_folder: str = None):
        self.upload_folder = upload_folder or config.UPLOAD_FOLDER
        self.validator = AttachmentValidator()
        
        # Ensure upload folder exists
        os.makedirs(self.upload_folder, exist_ok=True)
    
    async def save_uploaded_file(self, file_data: bytes, filename: str) -> Dict[str, Any]:
        """Save uploaded file and return metadata"""
        try:
            # Validate filename
            if not self.validator.validate_filename(filename):
                raise AttachmentError("Invalid filename")
            
            # Generate unique filename to prevent conflicts
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_hash = hashlib.md5(file_data).hexdigest()[:8]
            name, ext = os.path.splitext(filename)
            unique_filename = f"{name}_{timestamp}_{file_hash}{ext}"
            
            # Save file
            file_path = os.path.join(self.upload_folder, unique_filename)
            async with aiofiles.open(file_path, 'wb') as f:
                await f.write(file_data)
            
            # Validate saved file
            is_valid, error_msg = self.validator.validate_attachment(file_path, filename)
            if not is_valid:
                # Clean up invalid file
                os.remove(file_path)
                raise AttachmentError(error_msg)
            
            # Get file metadata
            file_size = len(file_data)
            mime_type, _ = mimetypes.guess_type(filename)
            
            return {
                'original_filename': filename,
                'stored_filename': unique_filename,
                'file_path': file_path,
                'file_size': file_size,
                'mime_type': mime_type,
                'uploaded_at': datetime.utcnow().isoformat(),
                'file_hash': hashlib.md5(file_data).hexdigest()
            }
        
        except Exception as e:
            logger.error(f"Error saving uploaded file: {e}")
            raise AttachmentError(f"Failed to save file: {str(e)}")
    
    async def process_attachment(self, file_path: str, filename: str) -> Dict[str, Any]:
        """Process existing file attachment"""
        try:
            # Validate file
            is_valid, error_msg = self.validator.validate_attachment(file_path, filename)
            if not is_valid:
                raise AttachmentError(error_msg)
            
            # Read file content
            async with aiofiles.open(file_path, 'rb') as f:
                content = await f.read()
            
            # Get file metadata
            file_size = len(content)
            mime_type, _ = mimetypes.guess_type(filename)
            
            return {
                'filename': filename,
                'content': content,
                'file_size': file_size,
                'mime_type': mime_type,
                'processed_at': datetime.utcnow().isoformat()
            }
        
        except Exception as e:
            logger.error(f"Error processing attachment: {e}")
            raise AttachmentError(f"Failed to process file: {str(e)}")
    
    def cleanup_file(self, file_path: str) -> bool:
        """Clean up temporary file"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        except Exception as e:
            logger.error(f"Error cleaning up file {file_path}: {e}")
            return False
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """Get file information"""
        try:
            if not os.path.exists(file_path):
                return {}
            
            stat = os.stat(file_path)
            mime_type, _ = mimetypes.guess_type(file_path)
            
            return {
                'filename': os.path.basename(file_path),
                'file_size': stat.st_size,
                'mime_type': mime_type,
                'created_at': datetime.fromtimestamp(stat.st_ctime).isoformat(),
                'modified_at': datetime.fromtimestamp(stat.st_mtime).isoformat()
            }
        except Exception as e:
            logger.error(f"Error getting file info: {e}")
            return {}

class AttachmentManager:
    """Manages attachments for email campaigns"""
    
    def __init__(self):
        self.processor = AttachmentProcessor()
        self.attachments = {}  # Store attachment metadata
    
    async def add_attachment(self, file_data: bytes, filename: str) -> str:
        """Add attachment and return attachment ID"""
        try:
            # Save file
            metadata = await self.processor.save_uploaded_file(file_data, filename)
            
            # Generate attachment ID
            attachment_id = hashlib.md5(
                f"{metadata['file_hash']}{metadata['uploaded_at']}".encode()
            ).hexdigest()
            
            # Store metadata
            self.attachments[attachment_id] = metadata
            
            return attachment_id
        
        except Exception as e:
            logger.error(f"Error adding attachment: {e}")
            raise AttachmentError(f"Failed to add attachment: {str(e)}")
    
    async def get_attachment(self, attachment_id: str) -> Dict[str, Any]:
        """Get attachment by ID"""
        try:
            if attachment_id not in self.attachments:
                raise AttachmentError("Attachment not found")
            
            metadata = self.attachments[attachment_id]
            
            # Process file
            processed = await self.processor.process_attachment(
                metadata['file_path'], 
                metadata['original_filename']
            )
            
            return {
                **metadata,
                **processed
            }
        
        except Exception as e:
            logger.error(f"Error getting attachment: {e}")
            raise AttachmentError(f"Failed to get attachment: {str(e)}")
    
    def remove_attachment(self, attachment_id: str) -> bool:
        """Remove attachment"""
        try:
            if attachment_id not in self.attachments:
                return False
            
            metadata = self.attachments[attachment_id]
            
            # Clean up file
            self.processor.cleanup_file(metadata['file_path'])
            
            # Remove from storage
            del self.attachments[attachment_id]
            
            return True
        
        except Exception as e:
            logger.error(f"Error removing attachment: {e}")
            return False
    
    def list_attachments(self) -> List[Dict[str, Any]]:
        """List all attachments"""
        return list(self.attachments.values())
    
    def get_attachment_info(self, attachment_id: str) -> Dict[str, Any]:
        """Get attachment metadata without loading content"""
        return self.attachments.get(attachment_id, {})
    
    def cleanup_old_attachments(self, max_age_hours: int = 24) -> int:
        """Clean up old attachments"""
        try:
            cutoff_time = datetime.utcnow().timestamp() - (max_age_hours * 3600)
            removed_count = 0
            
            for attachment_id, metadata in list(self.attachments.items()):
                uploaded_time = datetime.fromisoformat(metadata['uploaded_at']).timestamp()
                
                if uploaded_time < cutoff_time:
                    if self.remove_attachment(attachment_id):
                        removed_count += 1
            
            return removed_count
        
        except Exception as e:
            logger.error(f"Error cleaning up old attachments: {e}")
            return 0
    
    def get_total_size(self) -> int:
        """Get total size of all attachments"""
        total_size = 0
        for metadata in self.attachments.values():
            total_size += metadata.get('file_size', 0)
        return total_size
    
    def get_attachment_count(self) -> int:
        """Get number of attachments"""
        return len(self.attachments)

# Global attachment manager instance
attachment_manager = AttachmentManager()
