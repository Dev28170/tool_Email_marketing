"""
Dynamic timing utilities for email automation
Replaces static delays with condition-based waits for better performance and reliability
"""

import time
from typing import Optional, Callable, Any
from playwright.sync_api import Page, Locator
from loguru import logger


class DynamicTiming:
    """Dynamic timing utilities for Playwright automation"""
    
    # Class variables to track proxy performance
    _proxy_performance = {}
    _base_timeouts = {
        'page_load': 30000,
        'compose_load': 120000,  # Increased significantly for proxy loading issues
        'element_wait': 10000,  # Increased base timeout for element operations
        'network_idle': 10000,
        'send_confirmation': 10000,
        'attachment_upload': 5000  # Base delay for attachment operations
    }
    
    @staticmethod
    def get_proxy_timeout_multiplier(proxy: str = None) -> float:
        """Get timeout multiplier based on proxy performance history"""
        if not proxy:
            return 1.0
        
        # Get performance history for this proxy
        performance = DynamicTiming._proxy_performance.get(proxy, {
            'success_count': 0,
            'failure_count': 0,
            'avg_response_time': 0,
            'last_response_time': 0
        })
        
        # Calculate multiplier based on performance
        if performance['failure_count'] > 0:
            # Increase timeout for proxies with failures - more aggressive
            multiplier = 1.0 + (performance['failure_count'] * 1.0)  # Increased from 0.5 to 1.0
        elif performance['avg_response_time'] > 0:
            # Adjust based on average response time - more aggressive scaling
            if performance['avg_response_time'] > 15000:  # > 15 seconds
                multiplier = 3.0  # Increased from 2.0
            elif performance['avg_response_time'] > 10000:  # > 10 seconds
                multiplier = 2.5  # Increased from 2.0
            elif performance['avg_response_time'] > 5000:  # > 5 seconds
                multiplier = 2.0  # Increased from 1.5
            else:
                multiplier = 1.0
        else:
            multiplier = 1.0
        
        # Cap multiplier between 1.0 and 6.0 (increased for very slow proxies)
        return min(max(multiplier, 1.0), 6.0)
    
    @staticmethod
    def record_proxy_performance(proxy: str, success: bool, response_time: float):
        """Record proxy performance for future timeout adjustments"""
        if not proxy:
            return
        
        if proxy not in DynamicTiming._proxy_performance:
            DynamicTiming._proxy_performance[proxy] = {
                'success_count': 0,
                'failure_count': 0,
                'avg_response_time': 0,
                'last_response_time': 0
            }
        
        perf = DynamicTiming._proxy_performance[proxy]
        
        if success:
            perf['success_count'] += 1
        else:
            perf['failure_count'] += 1
        
        # Update average response time
        perf['last_response_time'] = response_time
        if perf['avg_response_time'] == 0:
            perf['avg_response_time'] = response_time
        else:
            # Exponential moving average
            perf['avg_response_time'] = (perf['avg_response_time'] * 0.8) + (response_time * 0.2)
    
    @staticmethod
    def get_adaptive_timeout(timeout_type: str, proxy: str = None) -> int:
        """Get adaptive timeout based on proxy performance"""
        base_timeout = DynamicTiming._base_timeouts.get(timeout_type, 5000)
        multiplier = DynamicTiming.get_proxy_timeout_multiplier(proxy)
        return int(base_timeout * multiplier)
    
    @staticmethod
    def wait_for_element_ready(page: Page, locator: Locator, max_wait: int = 5000, check_interval: int = 100, proxy: str = None) -> bool:
        """Wait for element to be ready (visible, enabled, stable) with dynamic timing"""
        start_time = time.time()
        while (time.time() - start_time) * 1000 < max_wait:
            try:
                if locator.is_visible(timeout=check_interval):
                    # Additional stability check
                    time.sleep(0.05)  # 50ms stability check
                    if locator.is_visible(timeout=check_interval):
                        return True
            except Exception:
                pass
            time.sleep(check_interval / 1000)
        return False
    
    @staticmethod
    def wait_for_network_idle(page: Page, max_wait: int = 10000, idle_time: int = 500, proxy: str = None) -> bool:
        """Wait for network to be idle with dynamic timing"""
        # Use adaptive timeout if proxy is provided
        if proxy:
            max_wait = DynamicTiming.get_adaptive_timeout('network_idle', proxy)
        
        try:
            page.wait_for_load_state('networkidle', timeout=max_wait)
            return True
        except Exception:
            # Fallback: wait for DOM to be stable
            try:
                page.wait_for_load_state('domcontentloaded', timeout=max_wait)
                time.sleep(idle_time / 1000)
                return True
            except Exception:
                return False
    
    @staticmethod
    def wait_for_condition(condition_func: Callable[[], bool], max_wait: int = 5000, check_interval: int = 100) -> bool:
        """Wait for a custom condition to be true with dynamic timing"""
        start_time = time.time()
        while (time.time() - start_time) * 1000 < max_wait:
            try:
                if condition_func():
                    return True
            except Exception:
                pass
            time.sleep(check_interval / 1000)
        return False
    
    @staticmethod
    def adaptive_delay(base_delay: int, success_count: int = 0, failure_count: int = 0) -> int:
        """Calculate adaptive delay based on success/failure history"""
        # Base delay with exponential backoff for failures
        if failure_count > 0:
            delay = base_delay * (1.5 ** min(failure_count, 3))  # Cap at 3x
        else:
            # Reduce delay for consecutive successes
            delay = max(base_delay * (0.8 ** min(success_count, 3)), base_delay // 2)
        
        return int(delay)
    
    @staticmethod
    def wait_for_bcc_chips_stable(page: Page, expected_count: int, max_wait: int = 3000) -> bool:
        """Wait for BCC chips to stabilize with dynamic timing"""
        def chips_stable():
            try:
                # Count current chips
                chip_selectors = [
                    "[data-automationid*='Pill']",
                    "[data-testid*='recipient-chip']", 
                    "[data-testid*='pill']",
                    "[role='button'][aria-label*='Remove']",
                    "[class*='persona']",
                    "[class*='chip']",
                    "[class*='pill']"
                ]
                
                total_chips = 0
                for selector in chip_selectors:
                    try:
                        count = page.locator(selector).count()
                        total_chips += count
                    except Exception:
                        continue
                
                return total_chips >= expected_count
            except Exception:
                return False
        
        return DynamicTiming.wait_for_condition(chips_stable, max_wait, 200)
    
    @staticmethod
    def wait_for_send_confirmation(page: Page, max_wait: int = 10000) -> bool:
        """Wait for send confirmation with comprehensive detection"""
        def send_confirmed():
            try:
                # Check for success indicators - more comprehensive list
                success_indicators = [
                    'text="Your message has been sent"',
                    'text="Message sent"',
                    'text="Email sent"',
                    'text="Sent"',
                    'text="Message delivered"',
                    '[data-testid*="success"]',
                    '.success-message',
                    '[aria-label*="sent"]',
                    '[aria-label*="delivered"]',
                    # Outlook-specific indicators
                    'div[role="alert"]:has-text("sent")',
                    'div[role="alert"]:has-text("delivered")',
                    'div[role="status"]:has-text("sent")',
                    'div[role="status"]:has-text("delivered")',
                    # Check for compose window disappearance
                    'div[aria-label*="New message"]:not([style*="display: none"])',
                    'div[role="dialog"][aria-label*="New message"]:not([style*="display: none"])'
                ]
                
                for indicator in success_indicators:
                    try:
                        if page.locator(indicator).is_visible(timeout=500):
                            return True
                    except Exception:
                        continue
                
                # Check if compose window is gone (more reliable than URL check)
                try:
                    compose_windows = page.locator('div[aria-label*="New message"], div[role="dialog"][aria-label*="New message"]').count()
                    if compose_windows == 0:
                        return True
                except Exception:
                    pass
                
                # Check if we're no longer on compose page
                if "compose" not in page.url.lower():
                    return True
                
                # Check if drafts folder shows the email is no longer there
                try:
                    # Look for the email in drafts - if it's gone, it was sent
                    draft_emails = page.locator('[data-testid*="draft"], [aria-label*="Draft"]').count()
                    if draft_emails == 0:
                        return True
                except Exception:
                    pass
                    
                return False
            except Exception:
                return False
        
        return DynamicTiming.wait_for_condition(send_confirmed, max_wait, 500)
    
    @staticmethod
    def wait_for_page_load(page: Page, url_pattern: str = None, max_wait: int = 15000, proxy: str = None) -> bool:
        """Wait for page to load with dynamic timing"""
        # Use adaptive timeout if proxy is provided
        if proxy:
            max_wait = DynamicTiming.get_adaptive_timeout('page_load', proxy)
        
        try:
            # Wait for network idle first
            if DynamicTiming.wait_for_network_idle(page, max_wait // 2, proxy=proxy):
                # Additional wait for specific URL pattern if provided
                if url_pattern:
                    def url_matches():
                        try:
                            return url_pattern.lower() in page.url.lower()
                        except Exception:
                            return False
                    
                    return DynamicTiming.wait_for_condition(url_matches, max_wait // 2, 200)
                return True
        except Exception:
            pass
        
        return False
    
    @staticmethod
    def smart_typing_delay(text_length: int, base_delay: int = 50) -> int:
        """Calculate typing delay based on text length"""
        if text_length <= 10:
            return base_delay
        elif text_length <= 50:
            return max(base_delay // 2, 20)
        else:
            return max(base_delay // 4, 10)


class TimingContext:
    """Context manager for tracking timing performance"""
    
    def __init__(self, operation_name: str):
        self.operation_name = operation_name
        self.start_time = None
        self.success_count = 0
        self.failure_count = 0
    
    def __enter__(self):
        self.start_time = time.time()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = (time.time() - self.start_time) * 1000
        if exc_type is None:
            self.success_count += 1
            logger.debug(f"✅ {self.operation_name} completed in {duration:.1f}ms")
        else:
            self.failure_count += 1
            logger.warning(f"❌ {self.operation_name} failed after {duration:.1f}ms: {exc_val}")
    
    def get_adaptive_delay(self, base_delay: int) -> int:
        """Get adaptive delay based on operation history"""
        return DynamicTiming.adaptive_delay(base_delay, self.success_count, self.failure_count)
    
    @staticmethod
    def get_attachment_delay(proxy: str = None) -> int:
        """Get adaptive delay for attachment operations"""
        base_delay = DynamicTiming._base_timeouts.get('attachment_upload', 5000)
        multiplier = DynamicTiming.get_proxy_timeout_multiplier(proxy)
        return int(base_delay * multiplier)
