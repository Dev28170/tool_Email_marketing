"""
HTML Email Processing Utilities
Handles HTML email content validation, sanitization, and processing
"""

import re
import html
from typing import Dict, List, Optional, Tuple
from bs4 import BeautifulSoup
import bleach
from loguru import logger


class HTMLEmailProcessor:
    """Professional HTML email content processor"""
    
    # Allowed HTML tags for email content
    ALLOWED_TAGS = [
        'p', 'br', 'strong', 'b', 'em', 'i', 'u', 'span', 'div',
        'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
        'a', 'img', 'table', 'tr', 'td', 'th', 'tbody', 'thead',
        'ul', 'ol', 'li', 'blockquote', 'hr',
        'font', 'center', 'pre', 'code'
    ]
    
    # Allowed attributes for specific tags
    ALLOWED_ATTRIBUTES = {
        'a': ['href', 'title', 'target', 'rel', 'style'],
        'img': ['src', 'alt', 'title', 'width', 'height', 'style'],
        'font': ['color', 'size', 'face'],
        'span': ['style', 'class'],
        'div': ['style', 'class', 'align'],
        'p': ['style', 'class', 'align'],
        'table': ['style', 'class', 'border', 'cellpadding', 'cellspacing', 'width'],
        'td': ['style', 'class', 'colspan', 'rowspan', 'align', 'valign'],
        'th': ['style', 'class', 'colspan', 'rowspan', 'align', 'valign'],
        'tr': ['style', 'class'],
        'tbody': ['style', 'class'],
        'thead': ['style', 'class'],
        'ul': ['style', 'class'],
        'ol': ['style', 'class'],
        'li': ['style', 'class'],
        'h1': ['style', 'class'],
        'h2': ['style', 'class'],
        'h3': ['style', 'class'],
        'h4': ['style', 'class'],
        'h5': ['style', 'class'],
        'h6': ['style', 'class'],
        'blockquote': ['style', 'class'],
        'hr': ['style', 'class'],
        'pre': ['style', 'class'],
        'code': ['style', 'class'],
        'center': ['style', 'class']
    }
    
    # Allowed CSS properties for style attributes
    ALLOWED_CSS_PROPERTIES = [
        'color', 'background-color', 'font-size', 'font-family', 'font-weight',
        'text-align', 'text-decoration', 'margin', 'padding', 'border',
        'width', 'height', 'display', 'line-height', 'letter-spacing',
        'word-spacing', 'text-indent', 'vertical-align', 'border-radius',
        'border-collapse', 'border-spacing', 'max-width', 'min-width',
        'background', 'background-image', 'background-position', 'background-repeat',
        'box-shadow', 'text-shadow', 'opacity', 'visibility', 'overflow',
        'white-space', 'word-wrap', 'text-overflow', 'cursor', 'outline',
        'border-top', 'border-right', 'border-bottom', 'border-left',
        'margin-top', 'margin-right', 'margin-bottom', 'margin-left',
        'padding-top', 'padding-right', 'padding-bottom', 'padding-left'
    ]
    
    def __init__(self):
        """Initialize HTML email processor"""
        self.soup = None
    
    def validate_html_content(self, html_content: str) -> Dict[str, any]:
        """
        Validate HTML content for email compatibility
        
        Args:
            html_content: Raw HTML content string
            
        Returns:
            Dict with validation results
        """
        result = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'sanitized_html': '',
            'has_html': False,
            'tag_count': 0,
            'allowed_tags': [],
            'disallowed_tags': []
        }
        
        try:
            if not html_content or not html_content.strip():
                result['warnings'].append('Empty HTML content')
                return result
            
            # Check if content contains HTML tags
            html_pattern = r'<[^>]+>'
            html_tags = re.findall(html_pattern, html_content)
            result['has_html'] = len(html_tags) > 0
            result['tag_count'] = len(html_tags)
            
            if not result['has_html']:
                result['warnings'].append('No HTML tags found - content will be treated as plain text')
                result['sanitized_html'] = html_content
                return result
            
            # Parse HTML with BeautifulSoup
            self.soup = BeautifulSoup(html_content, 'html.parser')
            
            # Find all tags
            all_tags = [tag.name for tag in self.soup.find_all()]
            unique_tags = list(set(all_tags))
            
            # Check for allowed/disallowed tags
            for tag in unique_tags:
                if tag in self.ALLOWED_TAGS:
                    result['allowed_tags'].append(tag)
                else:
                    result['disallowed_tags'].append(tag)
                    result['warnings'].append(f'Disallowed tag found: <{tag}>')
            
            # Check for potentially dangerous content
            dangerous_patterns = [
                r'javascript:', r'vbscript:', r'onload=', r'onclick=',
                r'onerror=', r'onmouseover=', r'<script', r'<iframe',
                r'<object', r'<embed', r'<form', r'<input'
            ]
            
            for pattern in dangerous_patterns:
                if re.search(pattern, html_content, re.IGNORECASE):
                    result['errors'].append(f'Potentially dangerous content found: {pattern}')
                    result['valid'] = False
            
            # Sanitize HTML content
            result['sanitized_html'] = self.sanitize_html(html_content)
            
            logger.info(f"HTML validation completed: {len(result['allowed_tags'])} allowed tags, {len(result['disallowed_tags'])} disallowed tags")
            
        except Exception as e:
            result['valid'] = False
            result['errors'].append(f'HTML validation error: {str(e)}')
            logger.error(f"HTML validation error: {e}")
        
        return result
    
    def sanitize_html(self, html_content: str) -> str:
        """
        Sanitize HTML content for safe email sending
        
        Args:
            html_content: Raw HTML content
            
        Returns:
            Sanitized HTML content
        """
        try:
            # Use bleach to sanitize HTML with CSS filtering
            from bleach.css_sanitizer import CSSSanitizer
            css_sanitizer = CSSSanitizer(allowed_css_properties=self.ALLOWED_CSS_PROPERTIES)
            
            sanitized = bleach.clean(
                html_content,
                tags=self.ALLOWED_TAGS,
                attributes=self.ALLOWED_ATTRIBUTES,
                css_sanitizer=css_sanitizer,
                strip=True,
                strip_comments=True
            )
            
            # Additional manual cleanup
            sanitized = self._manual_html_cleanup(sanitized)
            
            logger.info("HTML content sanitized successfully")
            return sanitized
            
        except Exception as e:
            logger.error(f"HTML sanitization error: {e}")
            # Fallback: escape HTML and return as plain text
            return html.escape(html_content)
    
    def _manual_html_cleanup(self, html_content: str) -> str:
        """Additional manual HTML cleanup"""
        try:
            # Remove empty tags
            html_content = re.sub(r'<(\w+)[^>]*>\s*</\1>', '', html_content)
            
            # Fix common HTML issues
            html_content = html_content.replace('&nbsp;', ' ')
            html_content = re.sub(r'\s+', ' ', html_content)  # Normalize whitespace
            
            # Ensure proper line breaks
            html_content = html_content.replace('<br>', '<br />')
            html_content = html_content.replace('<br/>', '<br />')
            
            return html_content.strip()
            
        except Exception as e:
            logger.warning(f"Manual HTML cleanup error: {e}")
            return html_content
    
    def convert_to_email_html(self, html_content: str) -> str:
        """
        Convert HTML content to email-compatible format
        
        Args:
            html_content: Sanitized HTML content
            
        Returns:
            Email-compatible HTML with proper structure
        """
        try:
            # If content doesn't have HTML structure, wrap it
            if not html_content.strip().startswith('<'):
                html_content = f'<div>{html_content}</div>'
            
            # Ensure proper HTML structure for email
            if not re.search(r'<html|<body|<div|<p', html_content, re.IGNORECASE):
                html_content = f'<div style="font-family: Arial, sans-serif; line-height: 1.6;">{html_content}</div>'
            
            # Add email-specific CSS inline styles
            html_content = self._inline_css_styles(html_content)
            
            return html_content
            
        except Exception as e:
            logger.error(f"Email HTML conversion error: {e}")
            return html_content
    
    def _inline_css_styles(self, html_content: str) -> str:
        """Convert CSS styles to inline styles for email compatibility"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')

            # Normalize anchor buttons: ensure padding/background/border-radius preserved
            for a in soup.find_all('a'):
                style = (a.get('style') or '')
                # Ensure display inline-block for button
                if 'display' not in style:
                    style += ' display:inline-block;'
                # Minimal default button styles if background-color present
                if 'background' in style or 'background-color' in style:
                    if 'padding' not in style:
                        style += ' padding:12px 18px;'
                    if 'border-radius' not in style:
                        style += ' border-radius:6px;'
                    if 'color' not in style:
                        style += ' color:#ffffff;'
                    if 'text-decoration' not in style:
                        style += ' text-decoration:none;'
                a['style'] = style.strip()

            # Common email-safe styles
            email_styles = {
                'p': 'margin: 0 0 10px 0;',
                'h1': 'font-size: 24px; font-weight: bold; margin: 0 0 15px 0;',
                'h2': 'font-size: 20px; font-weight: bold; margin: 0 0 12px 0;',
                'h3': 'font-size: 18px; font-weight: bold; margin: 0 0 10px 0;',
                'strong': 'font-weight: bold;',
                'em': 'font-style: italic;',
                'u': 'text-decoration: underline;',
                'a': 'color: #0066cc; text-decoration: underline;',
                'div': '',
            }

            for tag_name, css in email_styles.items():
                for tag in soup.find_all(tag_name):
                    tag['style'] = (tag.get('style') or '') + css

            return str(soup)
        except Exception as e:
            logger.warning(f"Inline CSS conversion error: {e}")
            return html_content
    
    def extract_plain_text(self, html_content: str) -> str:
        """
        Extract plain text from HTML content for text-only emails
        
        Args:
            html_content: HTML content
            
        Returns:
            Plain text version
        """
        try:
            if not html_content:
                return ''
            
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.decompose()
            
            # Get text and clean up
            text = soup.get_text()
            
            # Clean up whitespace
            lines = (line.strip() for line in text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            text = ' '.join(chunk for chunk in chunks if chunk)
            
            return text
            
        except Exception as e:
            logger.error(f"Plain text extraction error: {e}")
            return html_content
    
    def validate_email_html(self, html_content: str) -> Tuple[bool, str, List[str]]:
        """
        Comprehensive email HTML validation
        
        Args:
            html_content: HTML content to validate
            
        Returns:
            Tuple of (is_valid, sanitized_html, warnings)
        """
        validation_result = self.validate_html_content(html_content)
        
        if not validation_result['valid']:
            logger.error(f"HTML validation failed: {validation_result['errors']}")
            return False, html_content, validation_result['errors']
        
        # Convert to email-compatible format
        email_html = self.convert_to_email_html(validation_result['sanitized_html'])
        
        warnings = validation_result['warnings']
        if validation_result['disallowed_tags']:
            warnings.append(f"Removed disallowed tags: {', '.join(validation_result['disallowed_tags'])}")
        
        logger.info(f"Email HTML validation completed successfully")
        return True, email_html, warnings


# Global instance for easy access
html_email_processor = HTMLEmailProcessor()


def process_html_email_content(html_content: str) -> Dict[str, any]:
    """
    Main function to process HTML email content
    
    Args:
        html_content: Raw HTML content
        
    Returns:
        Dict with processed content and metadata
    """
    try:
        # Handle None input
        if html_content is None:
            html_content = ''
        
        is_valid, sanitized_html, warnings = html_email_processor.validate_email_html(html_content)
        
        # Extract plain text version
        plain_text = html_email_processor.extract_plain_text(sanitized_html)
        
        result = {
            'success': is_valid,
            'html_content': sanitized_html,
            'plain_text': plain_text,
            'warnings': warnings,
            'has_html': html_email_processor.validate_html_content(html_content)['has_html'],
            'original_length': len(html_content),
            'processed_length': len(sanitized_html)
        }
        
        if warnings:
            logger.warning(f"HTML processing warnings: {warnings}")
        
        return result
        
    except Exception as e:
        logger.error(f"HTML email processing error: {e}")
        return {
            'success': False,
            'html_content': html_content or '',
            'plain_text': html_content or '',
            'warnings': [f'Processing error: {str(e)}'],
            'has_html': False,
            'original_length': len(html_content) if html_content else 0,
            'processed_length': len(html_content) if html_content else 0
        }
