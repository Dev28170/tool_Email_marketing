"""
Email Processing Utilities
Handles email extraction, categorization, and validation
"""

import re
import requests
import time
from typing import List, Dict, Tuple
from loguru import logger
from email_validator import validate_email, EmailNotValidError
from config import config

class EmailProcessor:
    """Processes and categorizes email addresses"""
    
    def __init__(self, debounce_api_key: str = None):
        self.debounce_api_key = debounce_api_key or config.DEBOUNCE_API_KEY
        self.batch_size = config.DEBOUNCE_BATCH_SIZE
        self.timeout = config.DEBOUNCE_TIMEOUT
        self.office365_domains = {
            'outlook.com', 'hotmail.com', 'live.com', 'msn.com', 
            'office365.com', 'microsoft.com'
        }
        self.gsuite_domains = {
            'gmail.com', 'googlemail.com', 'google.com'
        }
    
    def extract_emails_from_text(self, text: str) -> List[str]:
        """Extract all email addresses from text content"""
        if not text:
            return []
        
        # Email regex pattern
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = re.findall(email_pattern, text, re.IGNORECASE)
        
        # Clean and validate emails
        valid_emails = []
        for email in emails:
            email = email.strip().lower()
            try:
                validate_email(email)
                valid_emails.append(email)
            except EmailNotValidError:
                logger.debug(f"Invalid email format: {email}")
                continue
        
        # Remove duplicates while preserving order
        seen = set()
        unique_emails = []
        for email in valid_emails:
            if email not in seen:
                seen.add(email)
                unique_emails.append(email)
        
        logger.info(f"Extracted {len(unique_emails)} valid emails from text")
        return unique_emails
    
    def categorize_emails(self, emails: List[str]) -> Dict[str, List[str]]:
        """Categorize emails by provider type"""
        office365_emails = []
        gsuite_emails = []
        other_emails = []
        
        for email in emails:
            domain = email.split('@')[1].lower()
            
            if domain in self.office365_domains:
                office365_emails.append(email)
            elif domain in self.gsuite_domains:
                gsuite_emails.append(email)
            else:
                other_emails.append(email)
        
        categorized = {
            'office365': office365_emails,
            'gsuite': gsuite_emails,
            'others': other_emails
        }
        
        logger.info(f"Categorized emails - Office365: {len(office365_emails)}, GSuite: {len(gsuite_emails)}, Others: {len(other_emails)}")
        return categorized
    
    def validate_with_debounce(self, emails: List[str], batch_size: int = None) -> List[str]:
        """Validate emails using Debounce.io API"""
        if not self.debounce_api_key:
            logger.warning("Debounce.io API key not provided, skipping validation")
            return emails
        
        if not emails:
            return []
        
        batch_size = batch_size or self.batch_size
        valid_emails = []
        
        # Process emails in batches
        for i in range(0, len(emails), batch_size):
            batch = emails[i:i + batch_size]
            logger.info(f"Validating batch {i//batch_size + 1}: {len(batch)} emails")
            
            try:
                # Debounce.io API call
                response = self._call_debounce_api(batch)
                if response:
                    batch_valid = self._parse_debounce_response(response, batch)
                    valid_emails.extend(batch_valid)
                    logger.info(f"Batch validation complete: {len(batch_valid)}/{len(batch)} valid")
                else:
                    logger.warning(f"Debounce API failed for batch, using original emails")
                    valid_emails.extend(batch)
                
                # Rate limiting - wait between batches
                if i + batch_size < len(emails):
                    time.sleep(1)
                    
            except Exception as e:
                logger.error(f"Debounce validation error for batch: {e}")
                valid_emails.extend(batch)  # Fallback to original emails
        
        logger.info(f"Debounce validation complete: {len(valid_emails)}/{len(emails)} emails valid")
        return valid_emails
    
    def _call_debounce_api(self, emails: List[str]) -> Dict:
        """Call Debounce.io API for email validation"""
        try:
            url = "https://api.debounce.io/v1"
            headers = {
                'Authorization': f'Bearer {self.debounce_api_key}',
                'Content-Type': 'application/json'
            }
            
            # Debounce.io expects emails in specific format
            data = {
                'emails': emails,
                'api_key': self.debounce_api_key
            }
            
            response = requests.post(url, json=data, headers=headers, timeout=self.timeout)
            response.raise_for_status()
            
            return response.json()
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Debounce API request failed: {e}")
            return None
        except Exception as e:
            logger.error(f"Debounce API error: {e}")
            return None
    
    def _parse_debounce_response(self, response: Dict, original_emails: List[str]) -> List[str]:
        """Parse Debounce.io API response and extract valid emails"""
        try:
            valid_emails = []
            
            # Debounce.io response structure may vary
            if 'result' in response:
                results = response['result']
                for email_data in results:
                    email = email_data.get('email', '').lower()
                    result = email_data.get('result', '').lower()
                    
                    # Accept deliverable and accept-all emails
                    if result in ['deliverable', 'accept-all', 'unknown']:
                        valid_emails.append(email)
            
            # Fallback: if response structure is different
            elif 'emails' in response:
                for email, status in response['emails'].items():
                    if status.lower() in ['deliverable', 'accept-all', 'unknown']:
                        valid_emails.append(email.lower())
            
            # If no valid structure found, return original emails
            if not valid_emails:
                logger.warning("Could not parse Debounce response, using original emails")
                return original_emails
            
            return valid_emails
            
        except Exception as e:
            logger.error(f"Error parsing Debounce response: {e}")
            return original_emails
    
    def process_export_file(self, file_content: str, account_email: str) -> Dict[str, str]:
        """Process exported contacts file and generate categorized files"""
        logger.info(f"Processing export file for account: {account_email}")
        
        # Extract emails from file content
        emails = self.extract_emails_from_text(file_content)
        
        if not emails:
            logger.warning("No emails found in export file")
            return {}
        
        # Categorize emails
        categorized = self.categorize_emails(emails)
        
        # Validate each category with Debounce.io
        processed_files = {}
        
        for category, email_list in categorized.items():
            if not email_list:
                continue
            
            logger.info(f"Processing {category} emails: {len(email_list)}")
            
            # Validate with Debounce.io
            valid_emails = self.validate_with_debounce(email_list)
            
            if valid_emails:
                # Generate file content
                file_content = '\n'.join(valid_emails)
                
                # Generate filename
                safe_email = account_email.replace('@', '_at_').replace('.', '_')
                filename = f"{safe_email}-{category}.txt"
                
                processed_files[category] = {
                    'filename': filename,
                    'content': file_content,
                    'count': len(valid_emails)
                }
                
                logger.info(f"Generated {category} file: {filename} with {len(valid_emails)} emails")
        
        return processed_files

# Global instance
email_processor = EmailProcessor()
