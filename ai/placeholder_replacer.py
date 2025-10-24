"""
AI-powered placeholder replacement system
Handles OpenAI integration for dynamic content generation
"""

import asyncio
import aiohttp
import re
from typing import Dict, Any, Optional, List
from datetime import datetime
from loguru import logger
from config import config

class PlaceholderReplacer:
    """AI-powered placeholder replacement system"""
    
    def __init__(self):
        self.openai_api_key = config.OPENAI_API_KEY
        self.model = config.OPENAI_MODEL
        self.max_tokens = config.OPENAI_MAX_TOKENS
        self.base_url = "https://api.openai.com/v1"
        
        # Cache for generated content
        self.cache = {}
        self.cache_ttl = 3600  # 1 hour
    
    def extract_placeholders(self, text: str) -> List[str]:
        """Extract all placeholders from text (case-insensitive).
        Returns placeholder names uppercased without brackets.
        """
        pattern = r'\[([a-zA-Z_][a-zA-Z0-9_]*)\]'
        return [m.upper() for m in re.findall(pattern, text or '')]
    
    def is_ai_prompt(self, text: str) -> bool:
        """Determine if text is an AI prompt or static text"""
        # Simple heuristic: if it ends with period or has more than 5 words, treat as prompt
        return text.strip().endswith('.') or len(text.split()) > 5
    
    async def generate_ai_content(self, prompt: str, context: Dict[str, Any] = None) -> str:
        """Generate content using OpenAI API"""
        if not self.openai_api_key:
            return f"(AI_KEY_NOT_PROVIDED) {prompt[:200]}"
        
        # Check cache first
        cache_key = f"{prompt}:{hash(str(context)) if context else ''}"
        if cache_key in self.cache:
            cached_data = self.cache[cache_key]
            if datetime.now().timestamp() - cached_data['timestamp'] < self.cache_ttl:
                return cached_data['content']
        
        try:
            # Prepare context for AI
            system_prompt = self._build_system_prompt(context)
            
            headers = {
                "Authorization": f"Bearer {self.openai_api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": self.max_tokens,
                "temperature": 0.7,
                "top_p": 1.0,
                "frequency_penalty": 0.0,
                "presence_penalty": 0.0
            }
            
            async with aiohttp.ClientSession() as session:
                async with session.post(
                    f"{self.base_url}/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=aiohttp.ClientTimeout(total=30)
                ) as response:
                    if response.status != 200:
                        error_text = await response.text()
                        logger.error(f"OpenAI API error: {response.status} - {error_text}")
                        return f"(OPENAI_ERROR {response.status}) {error_text[:200]}"
                    
                    data = await response.json()
                    
                    # Extract generated content
                    try:
                        content = data["choices"][0]["message"]["content"].strip()
                        
                        # Cache the result
                        self.cache[cache_key] = {
                            'content': content,
                            'timestamp': datetime.now().timestamp()
                        }
                        
                        return content
                    except (KeyError, IndexError) as e:
                        logger.error(f"Error parsing OpenAI response: {e}")
                        return "(OPENAI_PARSE_ERROR)"
        
        except asyncio.TimeoutError:
            logger.error("OpenAI API timeout")
            return "(OPENAI_TIMEOUT)"
        except Exception as e:
            logger.error(f"OpenAI API error: {e}")
            return f"(OPENAI_ERROR) {str(e)[:200]}"
    
    def _build_system_prompt(self, context: Dict[str, Any] = None) -> str:
        """Build system prompt with context"""
        base_prompt = """You are a professional email content generator. Generate concise, professional, and engaging content based on the user's request. Keep responses under 200 words and maintain a professional tone."""
        
        if context:
            context_info = []
            for key, value in context.items():
                if isinstance(value, (str, int, float)):
                    context_info.append(f"{key}: {value}")
                elif isinstance(value, dict):
                    context_info.append(f"{key}: {', '.join(f'{k}={v}' for k, v in value.items())}")
            
            if context_info:
                base_prompt += f"\n\nContext information:\n" + "\n".join(context_info)
        
        return base_prompt
    
    async def replace_placeholders(self, text: str, placeholder_data: Dict[str, Any] = None,
                                   *, batch_rand: Optional[str] = None,
                                   tz_name: str = "America/New_York",
                                   project_ai_prompt: Optional[str] = None) -> str:
        """Replace all placeholders in text with generated or static content.

        Special placeholders (case-insensitive):
          - [RAND]: replaced with 5 random digits; use provided batch_rand if given
          - [DATE]: current date/time formatted MM/DD/YYYY HH:MM in US/Eastern
          - [PROJECT_AI]: AI-generated 2 paragraphs about project bidding (per call)
        """
        if not text:
            return text
        
        # Extract all placeholders
        placeholders = self.extract_placeholders(text)
        if not placeholders:
            return text
        
        result = text
        
        from random import randint
        try:
            from zoneinfo import ZoneInfo  # Python 3.9+
        except Exception:
            ZoneInfo = None  # Fallback to naive time if zoneinfo not available

        for placeholder in placeholders:
            key = placeholder.upper()
            placeholder_key_patterns = [f"[{placeholder}]", f"[{placeholder.lower()}]", f"[{placeholder.upper()}]"]
            
            def replace_all(target: str):
                nonlocal result
                for pat in placeholder_key_patterns:
                    if pat in result:
                        result = result.replace(pat, target)

            # Support both [RAND] and [RAND5] -> 5 random digits
            if key in ('RAND', 'RAND5'):
                rand_value = batch_rand if batch_rand is not None else f"{randint(0, 99999):05d}"
                replace_all(rand_value)
                continue
            if key == 'DATE':
                now = datetime.now()
                try:
                    if ZoneInfo is not None:
                        now = datetime.now(ZoneInfo(tz_name))
                except Exception:
                    pass
                formatted = now.strftime('%m/%d/%Y %H:%M')
                replace_all(formatted)
                continue
            if key == 'PROJECT_AI':
                prompt = project_ai_prompt or (
                    "Write two short professional paragraphs discussing the project bidding process, "
                    "covering proposal preparation, evaluation criteria, competitiveness, timelines, and communication."
                )
                generated_content = await self.generate_ai_content(prompt, context=placeholder_data)
                replace_all(generated_content)
                continue
            
            # Generic placeholder behavior
            replacement_data = None
            if placeholder_data and key in placeholder_data:
                replacement_data = placeholder_data[key]
            elif placeholder_data and placeholder.lower() in placeholder_data:
                replacement_data = placeholder_data[placeholder.lower()]
            else:
                replacement_data = f"Please generate content for {key.lower().replace('_', ' ')}"
            
            if isinstance(replacement_data, str) and self.is_ai_prompt(replacement_data):
                generated_content = await self.generate_ai_content(
                    replacement_data,
                    context=placeholder_data
                )
                replace_all(generated_content)
            else:
                static_text = str(replacement_data) if replacement_data is not None else f"[{key}]"
                replace_all(static_text)
        
        return result
    
    async def replace_placeholders_batch(self, texts: List[str], 
                                       placeholder_data: Dict[str, Any] = None) -> List[str]:
        """Replace placeholders in multiple texts concurrently"""
        tasks = [
            self.replace_placeholders(text, placeholder_data) 
            for text in texts
        ]
        return await asyncio.gather(*tasks)
    
    def get_placeholder_suggestions(self, text: str) -> List[Dict[str, str]]:
        """Get suggestions for common placeholders"""
        suggestions = [
            {"placeholder": "PROJECT_DETAILS", "description": "Project status and updates"},
            {"placeholder": "CLIENT_NAME", "description": "Client or recipient name"},
            {"placeholder": "COMPANY_NAME", "description": "Company or organization name"},
            {"placeholder": "DEADLINE", "description": "Project deadline or important date"},
            {"placeholder": "BUDGET", "description": "Budget information or cost details"},
            {"placeholder": "TEAM_MEMBERS", "description": "Team member names or roles"},
            {"placeholder": "MILESTONE", "description": "Project milestone or achievement"},
            {"placeholder": "NEXT_STEPS", "description": "Next steps or action items"},
            {"placeholder": "CONTACT_INFO", "description": "Contact information"},
            {"placeholder": "MEETING_TIME", "description": "Meeting time or schedule"},
        ]
        
        # Filter suggestions based on existing placeholders
        existing_placeholders = self.extract_placeholders(text)
        return [s for s in suggestions if s["placeholder"] not in existing_placeholders]
    
    def validate_placeholders(self, text: str, placeholder_data: Dict[str, Any] = None) -> Dict[str, Any]:
        """Validate placeholder data and return status"""
        placeholders = self.extract_placeholders(text)
        
        validation = {
            'valid': True,
            'placeholders': placeholders,
            'missing_data': [],
            'ai_prompts': [],
            'static_texts': [],
            'warnings': []
        }
        
        for placeholder in placeholders:
            if placeholder_data and placeholder in placeholder_data:
                data = placeholder_data[placeholder]
                if isinstance(data, str) and self.is_ai_prompt(data):
                    validation['ai_prompts'].append(placeholder)
                else:
                    validation['static_texts'].append(placeholder)
            else:
                validation['missing_data'].append(placeholder)
                validation['valid'] = False
        
        # Check for AI key
        if validation['ai_prompts'] and not self.openai_api_key:
            validation['warnings'].append("OpenAI API key not configured - AI placeholders will show fallback text")
        
        return validation
    
    def clear_cache(self):
        """Clear the content cache"""
        self.cache.clear()
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """Get cache statistics"""
        now = datetime.now().timestamp()
        valid_entries = sum(1 for entry in self.cache.values() 
                          if now - entry['timestamp'] < self.cache_ttl)
        
        return {
            'total_entries': len(self.cache),
            'valid_entries': valid_entries,
            'expired_entries': len(self.cache) - valid_entries,
            'cache_ttl': self.cache_ttl
        }

    async def render_subject_body(self, subject: str, body: str,
                                  placeholder_data: Dict[str, Any] = None,
                                  *, per_batch_rand: Optional[str] = None,
                                  tz_name: str = "America/New_York") -> Dict[str, str]:
        """Render subject and body with placeholders.
        per_batch_rand allows keeping the same [RAND] within a batch while
        different across batches.
        """
        rendered_subject = await self.replace_placeholders(
            subject or '', placeholder_data, batch_rand=per_batch_rand, tz_name=tz_name
        )
        rendered_body = await self.replace_placeholders(
            body or '', placeholder_data, batch_rand=per_batch_rand, tz_name=tz_name
        )
        return { 'subject': rendered_subject, 'body': rendered_body }
    
    async def test_ai_connection(self) -> Dict[str, Any]:
        """Test OpenAI API connection"""
        if not self.openai_api_key:
            return {
                'success': False,
                'error': 'OpenAI API key not configured'
            }
        
        try:
            test_content = await self.generate_ai_content("Generate a simple test message.")
            
            if test_content.startswith("(OPENAI_ERROR") or test_content.startswith("(AI_KEY_NOT_PROVIDED)"):
                return {
                    'success': False,
                    'error': test_content
                }
            
            return {
                'success': True,
                'test_content': test_content,
                'model': self.model,
                'max_tokens': self.max_tokens
            }
        
        except Exception as e:
            return {
                'success': False,
                'error': f"Connection test failed: {str(e)}"
            }

# Global instance
placeholder_replacer = PlaceholderReplacer()
