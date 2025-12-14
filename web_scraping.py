from urllib.parse import urljoin, urlparse
import requests
import json
import re
from typing import Dict, Optional, List, Set, Any
from dataclasses import dataclass, asdict
from openai import AzureOpenAI
import logging
from datetime import datetime
import time
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd

import threading
from collections import defaultdict
import uuid

import requests
from requests.adapters import HTTPAdapter
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
from urllib3.util.retry import Retry
import atexit
import random


import resource
import sys

# Get current limits
soft, hard = resource.getrlimit(resource.RLIMIT_NOFILE)
print(f"Current file descriptor limits - Soft: {soft}, Hard: {hard}")

# Try to increase to 4096
try:
    resource.setrlimit(resource.RLIMIT_NOFILE, (8192, 8192))
    print("Successfully increased file descriptor limit to 4096")
except:
    # If we can't set to 4096, try to use the maximum allowed
    try:
        resource.setrlimit(resource.RLIMIT_NOFILE, (hard, hard))
        print(f"Set file descriptor limit to maximum allowed: {hard}")
    except:
        print(f"Could not change file descriptor limit, staying at: {soft}")

# Verify final limit
soft, hard = resource.getrlimit(resource.RLIMIT_NOFILE)
print(f"Final file descriptor limit: {soft}")


import urllib3

# Configure urllib3 to handle high concurrency
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)



logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SessionPool:
    """Thread-safe session pool with proper resource management"""
    _local = threading.local()
    _sessions = []
    _lock = threading.Lock()
    
    @classmethod
    def get_session(cls):
        if not hasattr(cls._local, 'session'):
            session = requests.Session()
            
            # Configure aggressive connection pooling for high concurrency
            adapter = HTTPAdapter(
                pool_connections=100,  # Increased for your 150 workers
                pool_maxsize=200,      # Increased for your workload
                max_retries=Retry(
                    total=3, 
                    backoff_factor=1,
                    status_forcelist=[500, 502, 503, 504]
                ),
                pool_block=False
            )
            session.mount('http://', adapter)
            session.mount('https://', adapter)
            
            # Set connection timeouts
            session.timeout = 120
            
            cls._local.session = session
            
            # Track sessions for cleanup
            with cls._lock:
                cls._sessions.append(session)
        
        return cls._local.session
    
    @classmethod
    def close_all_sessions(cls):
        """Close all sessions - call at program end"""
        with cls._lock:
            for session in cls._sessions:
                try:
                    session.close()
                except:
                    pass
            cls._sessions.clear()

# Global session pool
session_pool = SessionPool()

# Register cleanup at exit
atexit.register(session_pool.close_all_sessions)

class BatchLogger:
    """Thread-safe logger for batch processing"""
    def __init__(self):
        self.logs = defaultdict(dict)  # Change from defaultdict(list) to defaultdict(dict)
        self.lock = threading.Lock()
    
    def log(self, batch_id: str, url: str, message: str, level: str = "INFO"):
        with self.lock:
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "level": level,
                "message": message
            }
            if batch_id not in self.logs:
                self.logs[batch_id] = {}
            if url not in self.logs[batch_id]:
                self.logs[batch_id][url] = {"logs": [], "child_urls": {}}
            self.logs[batch_id][url]["logs"].append(log_entry)
    
    def log_child(self, batch_id: str, parent_url: str, child_url: str, message: str, level: str = "INFO"):
        with self.lock:
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "level": level,
                "message": message
            }
            if batch_id not in self.logs:
                self.logs[batch_id] = {}
            if parent_url not in self.logs[batch_id]:
                self.logs[batch_id][parent_url] = {"logs": [], "child_urls": {}}
            if child_url not in self.logs[batch_id][parent_url]["child_urls"]:
                self.logs[batch_id][parent_url]["child_urls"][child_url] = []
            self.logs[batch_id][parent_url]["child_urls"][child_url].append(log_entry)
    
    def get_logs(self, batch_id: str):
        with self.lock:
            return self.logs.get(batch_id, {})

# Global batch logger instance
batch_logger = BatchLogger()

PROXY_COUNTRIES = [
    'UnitedStates', 
    'Canada', 
    'UnitedKingdom', 
    'Germany', 
    'France', 
    'Netherlands',
    'Australia',
    'Singapore'
]

class ProxyRotator:
    """Manages proxy country rotation with domain-specific tracking"""
    
    def __init__(self):
        self.domain_proxy_success = defaultdict(lambda: defaultdict(int))
        self.domain_proxy_failures = defaultdict(lambda: defaultdict(int))
        self.lock = threading.Lock()
        self.rotation_index = 0
    
    def get_next_proxy(self, domain: str = None, attempt: int = 0) -> str:
        """Get next proxy country based on attempt number and domain history"""
        with self.lock:
            if domain and attempt == 0:
                # First attempt: use best performing proxy for this domain
                if domain in self.domain_proxy_success:
                    proxy_scores = {}
                    for proxy in PROXY_COUNTRIES:
                        success = self.domain_proxy_success[domain][proxy]
                        failures = self.domain_proxy_failures[domain][proxy]
                        total = success + failures
                        if total > 0:
                            proxy_scores[proxy] = success / total
                    
                    if proxy_scores:
                        best_proxy = max(proxy_scores, key=proxy_scores.get)
                        logger.info(f"Using best proxy for {domain}: {best_proxy}")
                        return best_proxy
            
            # Round-robin with attempt offset
            index = (self.rotation_index + attempt) % len(PROXY_COUNTRIES)
            self.rotation_index += 1
            return PROXY_COUNTRIES[index]
    
    def record_success(self, domain: str, proxy: str):
        """Record successful scrape with proxy"""
        with self.lock:
            self.domain_proxy_success[domain][proxy] += 1
    
    def record_failure(self, domain: str, proxy: str):
        """Record failed scrape with proxy"""
        with self.lock:
            self.domain_proxy_failures[domain][proxy] += 1


class SessionManager:
    """Manages persistent sessions per domain"""
    
    def __init__(self):
        self.domain_sessions = {}
        self.lock = threading.Lock()
    
    def get_session_id(self, domain: str) -> str:
        """Get or create session ID for domain"""
        with self.lock:
            if domain not in self.domain_sessions:
                self.domain_sessions[domain] = str(uuid.uuid4())
                logger.info(f"Created new session for {domain}: {self.domain_sessions[domain]}")
            return self.domain_sessions[domain]
    
    def clear_session(self, domain: str):
        """Clear session for domain (useful after failures)"""
        with self.lock:
            if domain in self.domain_sessions:
                del self.domain_sessions[domain]
                logger.info(f"Cleared session for {domain}")


class BrowserActionsBuilder:
    """Builds realistic browser action sequences"""
    
    @staticmethod
    def get_basic_actions() -> List[Dict]:
        """Basic wait and scroll"""
        return [
            {'type': 'wait', 'wait': random.randint(2, 4), 'when': 'afterload'},
            {'type': 'scroll', 'x': 0, 'y': random.randint(300, 600), 'when': 'afterload'}
        ]
    
    @staticmethod
    def get_cookie_consent_actions() -> List[Dict]:
        """Actions to handle cookie consent banners"""
        return [
            {'type': 'wait', 'wait': 2, 'when': 'afterload'},
            # Try multiple common cookie consent selectors
            {'type': 'click', 'selector': 'button[id*="accept"]', 'when': 'afterload'},
            {'type': 'click', 'selector': '.cookie-accept', 'when': 'afterload'},
            {'type': 'click', 'selector': '#accept-cookies', 'when': 'afterload'},
            {'type': 'click', 'selector': '[class*="cookie"][class*="accept"]', 'when': 'afterload'},
            {'type': 'click', 'selector': 'button[class*="consent"]', 'when': 'afterload'},
            {'type': 'wait', 'wait': 1, 'when': 'afterload'},
            {'type': 'scroll', 'x': 0, 'y': 500, 'when': 'afterload'}
        ]
    
    @staticmethod
    def get_anti_bot_actions() -> List[Dict]:
        """Aggressive actions for anti-bot protection"""
        return [
            {'type': 'wait', 'wait': 5, 'when': 'afterload'},
            {'type': 'scroll', 'x': 0, 'y': 300, 'when': 'afterload'},
            {'type': 'wait', 'wait': 2, 'when': 'afterload'},
            {'type': 'scroll', 'x': 0, 'y': 600, 'when': 'afterload'},
            {'type': 'wait', 'wait': 2, 'when': 'afterload'},
            {'type': 'scroll', 'x': 0, 'y': 1000, 'when': 'afterload'},
            {'type': 'wait', 'wait': 3, 'when': 'afterload'}
        ]
    
    @staticmethod
    def get_actions_for_attempt(attempt: int, has_issues: bool = False) -> List[Dict]:
        """Get appropriate actions based on retry attempt"""
        if attempt == 0:
            # First attempt: basic
            return BrowserActionsBuilder.get_basic_actions()
        elif attempt == 1:
            # Second attempt: handle cookies
            return BrowserActionsBuilder.get_cookie_consent_actions()
        else:
            # Third+ attempt: full anti-bot
            return BrowserActionsBuilder.get_anti_bot_actions()




# Update AdaptiveCrawler class - ADD THESE METHODS
class AdaptiveCrawler:
    """Track problematic sites and adapt strategy"""
    
    def __init__(self):
        self.site_issues = {}  # Track issues per domain
        self.lock = threading.Lock()
        self.domain_strategies = {}  # Store successful strategies
    
    def record_issue(self, url: str, issue_type: str):
        """Record an issue for a domain"""
        domain = urlparse(url).netloc
        with self.lock:
            if domain not in self.site_issues:
                self.site_issues[domain] = {'timeouts': 0, 'errors': 0, 'bot_detected': 0}
            
            if issue_type == 'timeout':
                self.site_issues[domain]['timeouts'] += 1
            elif issue_type == 'bot_detected':
                self.site_issues[domain]['bot_detected'] += 1
            else:
                self.site_issues[domain]['errors'] += 1
    
    def should_use_enhanced_request(self, url: str) -> bool:
        """Check if a domain needs enhanced handling"""
        domain = urlparse(url).netloc
        with self.lock:
            if domain in self.site_issues:
                issues = self.site_issues[domain]
                return issues['timeouts'] > 0 or issues['errors'] > 1 or issues['bot_detected'] > 0
        return False
    
    def get_issue_level(self, url: str) -> str:
        """Get severity level of issues for domain"""
        domain = urlparse(url).netloc
        with self.lock:
            if domain not in self.site_issues:
                return 'none'
            
            issues = self.site_issues[domain]
            if issues['bot_detected'] > 0:
                return 'high'
            elif issues['timeouts'] > 1 or issues['errors'] > 2:
                return 'medium'
            elif issues['timeouts'] > 0 or issues['errors'] > 0:
                return 'low'
        return 'none'
    
    def record_success(self, url: str, strategy: Dict):
        """Record successful strategy for domain"""
        domain = urlparse(url).netloc
        with self.lock:
            self.domain_strategies[domain] = strategy


# Create global instances - ADD THESE THREE LINES TOGETHER
proxy_rotator = ProxyRotator()
session_manager = SessionManager()
adaptive_crawler = AdaptiveCrawler()

# Environment variables are loaded from .env file
# Ensure AZURE_OPENAI_API_KEY, AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_VERSION, and AZURE_OPENAI_DEPLOYMENT_NAME are set in .env



def read_ria_urls_from_excel(file_path: str, url_column: str = 'Website Address') -> List[str]:
    """Read RIA URLs from Excel file with improved URL cleaning"""
    df = pd.read_excel(file_path, engine='openpyxl')
    urls = df[url_column].dropna().tolist()
    # Clean URLs
    cleaned_urls = []
    for url in urls:
        if pd.notna(url) and url.strip():
            url = url.strip()
            # Convert to lowercase for protocol only, preserve case for domain
            if url.lower().startswith('https://'):
                url = 'https://' + url[8:]
            elif url.lower().startswith('http://'):
                url = 'http://' + url[7:]
            else:
                # Remove any protocol-like prefix
                url = re.sub(r'^[Hh][Tt][Tt][Pp][Ss]?://', '', url)
                url = 'https://' + url
            
            # Ensure no double slashes after protocol
            url = re.sub(r'https://+', 'https://', url)
            url = re.sub(r'http://+', 'http://', url)
            
            cleaned_urls.append(url)
    return cleaned_urls

def clean_single_url(url: str) -> str:
    """Clean a single URL to ensure proper format"""
    if not url:
        return url
    
    url = url.strip()
    
    # Handle case-insensitive protocol
    if url.lower().startswith('https://'):
        url = 'https://' + url[8:]
    elif url.lower().startswith('http://'):
        url = 'http://' + url[7:]
    else:
        # Remove any malformed protocol
        url = re.sub(r'^[Hh][Tt][Tt][Pp][Ss]?://', '', url)
        url = 'https://' + url
    
    # Ensure no double slashes after protocol
    url = re.sub(r'https://+', 'https://', url)
    
    return url

@dataclass(frozen=True)
class CrawlConfig:
    """Crawler configuration"""
    api_key: str
    azure_openai_api_key: str  # Changed from openai_api_key
    azure_openai_endpoint: str  # ADD THIS
    azure_openai_api_version: str  # ADD THIS
    azure_openai_deployment: str  # ADD THIS
    base_url: str = 'https://publisher.scrappey.com/api/v1'
    max_pages: int = 20
    delay: float = 1.0
    timeout: int = 180

@dataclass
class RIAInfo:
    """Structured RIA information extracted from website"""
    company_name: str = ""
    services: List[str] = None
    target_clients: List[str] = None
    investment_strategy: str = ""
    products_instruments: List[str] = None
    key_personnel: List[Dict[str, str]] = None
    industry_insights: str = ""
    contact_info: Dict[str, str] = None
    website_url: str = ""
    scraped_at: str = ""
    pages_analyzed: int = 0
    all_page_data: List[Dict] = None

def create_scrape_request(config: CrawlConfig, url: str, attempt: int = 0) -> Dict:
    """Create scraping request with adaptive anti-bot strategies"""
    
    domain = urlparse(url).netloc
    issue_level = adaptive_crawler.get_issue_level(url)
    
    # Get appropriate proxy
    proxy_country = proxy_rotator.get_next_proxy(domain, attempt)
    
    # Get session ID for persistent session
    session_id = session_manager.get_session_id(domain)
    
    # Base request
    request_params = {
        'headers': {'Content-Type': 'application/json'},
        'params': {'key': config.api_key},
        'json_data': {
            'cmd': 'request.get',
            'url': url,
            'proxyCountry': proxy_country,
          #  "premiumProxy": True,
            'requestType': 'browser',
            'includeLinks': True,
            'includeImages': False,
            'sessionId': session_id  # Maintain session
        }
    }
    
    # Add localStorage for consent (always helpful)
    request_params['json_data']['localStorage'] = {
        'consent': 'true',
        'cookie_consent': 'accepted',
        'gdpr_consent': 'true'
    }
    
    # Add sessionStorage
    request_params['json_data']['sessionStorage'] = {
        'visited': 'true'
    }
    
    # Adaptive browser actions based on attempt and issue level
    if attempt == 0 and issue_level == 'none':
        # First attempt, no known issues: basic actions
        request_params['json_data']['browserActions'] = BrowserActionsBuilder.get_basic_actions()
        logger.info(f"Using basic actions for {domain} (attempt {attempt})")
    
    elif attempt == 1 or issue_level == 'low':
        # Second attempt or minor issues: cookie handling
        request_params['json_data']['browserActions'] = BrowserActionsBuilder.get_cookie_consent_actions()
        logger.info(f"Using cookie consent actions for {domain} (attempt {attempt})")
    
    elif attempt >= 2 or issue_level in ['medium', 'high']:
        # Multiple attempts or serious issues: full anti-bot
        request_params['json_data']['browserActions'] = BrowserActionsBuilder.get_anti_bot_actions()
        request_params['json_data']['datacenter'] = True  # Switch to datacenter proxy
        logger.info(f"Using anti-bot actions + datacenter proxy for {domain} (attempt {attempt})")
    
    # For very problematic sites, add extra measures
    if issue_level == 'high' or attempt >= 3:
        request_params['json_data']['browserActions'].append(
            {'type': 'wait', 'wait': 10, 'when': 'afterload'}
        )
        logger.warning(f"High issue level for {domain} - using maximum wait times")
    
    logger.info(f"Request config - Domain: {domain}, Proxy: {proxy_country}, Attempt: {attempt}, Issue Level: {issue_level}")
    
    return request_params

def execute_api_call(config: CrawlConfig, request_data: Dict, max_retries: int = 4) -> Dict:  # Changed to 4 retries
    """Execute API call with proxy rotation and adaptive retry logic"""
    
    url = request_data['json_data'].get('url', 'Unknown')
    domain = urlparse(url).netloc
    session = session_pool.get_session()
    
    # Make a deep copy to avoid modifying the original
    request_data = {
        'headers': request_data['headers'].copy(),
        'params': request_data['params'].copy(),
        'json_data': request_data['json_data'].copy()
    }
    
    for attempt in range(max_retries):
        response = None
        current_proxy = request_data['json_data'].get('proxyCountry', 'UnitedStates')
        
        try:
            # Increase timeout progressively
            current_timeout = 60 + (attempt * 30)  # 60, 90, 120, 150
            
            logger.info(f"Attempt {attempt + 1}/{max_retries}: Scraping {url} with proxy {current_proxy}")
            
            response = session.post(
                config.base_url,
                params=request_data['params'],
                headers=request_data['headers'],
                json=request_data['json_data'],
                timeout=current_timeout
            )
            
            logger.info(f"Response status code: {response.status_code}")
            response.raise_for_status()
            result = response.json()
            
            # Check for Scrappey error codes first
            if 'error' in result:
                error_code = result.get('code', 'UNKNOWN')
                error_msg = result.get('error', 'Unknown error')
                logger.error(f"Scrappey Error {error_code}: {error_msg}")
                
                # Record failure with this proxy
                proxy_rotator.record_failure(domain, current_proxy)
                
                # Categorize errors
                proxy_errors = ['CODE-0024', 'CODE-0025', 'CODE-0007', 'CODE-0019']
                temporary_errors = ['CODE-0001', 'CODE-0017', 'CODE-0031']
                blocking_errors = ['CODE-0002', 'CODE-0003', 'CODE-0008', 'CODE-0010', 'CODE-0037']
                non_retryable = ['CODE-0004', 'CODE-0018', 'CODE-0029', 'CODE-0030']
                browser_errors = ['CODE-0026', 'CODE-0005', 'CODE-0006']
                
                # SSL/Network errors
                ssl_network_indicators = [
                    'SSL_ERROR', 'NS_ERROR_NET_INTERRUPT', 'certificate', 'SSL', 'NET_INTERRUPT'
                ]
                
                if any(indicator in error_msg for indicator in ssl_network_indicators):
                    logger.error(f"SSL/Network error - website issue, not retrying")
                    return {}
                
                # Non-retryable errors
                if error_code in non_retryable:
                    logger.error(f"Non-retryable error - config/limit issue")
                    return {}
                
                # Anti-bot detection
                if error_code in blocking_errors:
                    logger.warning(f"Anti-bot protection detected!")
                    adaptive_crawler.record_issue(url, 'bot_detected')
                    
                    if attempt < max_retries - 1:
                        # Recreate request with enhanced strategy
                        request_data = create_scrape_request(config, url, attempt + 1)
                        time.sleep(30)  # Longer wait for bot protection
                        continue
                    return {}
                
                # Proxy issues - switch proxy
                if error_code in proxy_errors:
                    logger.warning(f"Proxy issue - switching proxy")
                    if attempt < max_retries - 1:
                        # Get new proxy and recreate request
                        request_data = create_scrape_request(config, url, attempt + 1)
                        time.sleep(20)
                        continue
                    return {}
                
                # Temporary/browser errors
                if error_code in temporary_errors or error_code in browser_errors:
                    logger.warning(f"Temporary/browser error - retrying with new strategy")
                    if attempt < max_retries - 1:
                        request_data = create_scrape_request(config, url, attempt + 1)
                        time.sleep(15)
                        continue
                    return {}
                
                # Unknown error - try different approach
                if attempt < max_retries - 1:
                    request_data = create_scrape_request(config, url, attempt + 1)
                    time.sleep(15)
                    continue
                return {}
            
            # Check if request was successful
            if result.get('data') != 'success':
                logger.warning(f"API returned non-success status: {result.get('data')}")
                if attempt < max_retries - 1:
                    request_data = create_scrape_request(config, url, attempt + 1)
                    time.sleep(10)
                    continue
                return {}
            
            # Validate content
            solution = result.get('solution', {})
            inner_text = solution.get('innerText', '').strip()
            html_response = solution.get('response', '').strip()
            
            logger.info(f"Content received - innerText: {len(inner_text)} chars, HTML: {len(html_response)} chars")
            
            # Check if content is actually present
            if not inner_text and not html_response:
                logger.warning(f"Empty content received for {url} on attempt {attempt + 1}")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying with enhanced strategy...")
                    request_data = create_scrape_request(config, url, attempt + 1)
                    time.sleep(10 * (attempt + 1))
                    continue
                else:
                    logger.error(f"Failed to get content after {max_retries} attempts")
                    return {}
            
            # Check for blocking patterns
            blocking_indicators = [
                'cloudflare', 'captcha', 'access denied', '403 forbidden', 
                'please verify you are human', 'checking your browser',
                'security check', 'ddos protection', 'ray id'
            ]
            
            combined_content = (inner_text + html_response).lower()
            for indicator in blocking_indicators:
                if indicator in combined_content:
                    logger.warning(f"Detected blocking pattern: {indicator}")
                    adaptive_crawler.record_issue(url, 'bot_detected')
                    
                    if attempt < max_retries - 1:
                        logger.info(f"Retrying with anti-bot strategy...")
                        request_data = create_scrape_request(config, url, attempt + 1)
                        time.sleep(20)
                        continue
                    break
            
            # If we have valid content, record success and return
            if len(inner_text) > 100 or len(html_response) > 500:
                logger.info(f"✅ Successfully scraped {url} on attempt {attempt + 1} with proxy {current_proxy}")
                
                # Record success
                proxy_rotator.record_success(domain, current_proxy)
                adaptive_crawler.record_success(url, {
                    'proxy': current_proxy,
                    'attempt': attempt,
                    'browser_actions': request_data['json_data'].get('browserActions', [])
                })
                
                return result
            else:
                logger.warning(f"Minimal content found ({len(inner_text)} chars text, {len(html_response)} chars HTML)")
                if attempt < max_retries - 1:
                    request_data = create_scrape_request(config, url, attempt + 1)
                    time.sleep(10)
                    continue
                    
            return result
            
        except requests.exceptions.Timeout:
            logger.warning(f"Timeout on attempt {attempt + 1}/{max_retries} for URL: {url}")
            proxy_rotator.record_failure(domain, current_proxy)
            
            if attempt < max_retries - 1:
                # Try with different proxy
                request_data = create_scrape_request(config, url, attempt + 1)
                time.sleep(15 * (attempt + 1))
                continue
            else:
                logger.error(f"Final timeout after {max_retries} attempts for {url}")
                adaptive_crawler.record_issue(url, 'timeout')
                return {}
                
        except requests.exceptions.HTTPError as e:
            error_text = e.response.text if hasattr(e, 'response') and e.response else str(e)
            proxy_rotator.record_failure(domain, current_proxy)
            
            # Special handling for 500 errors
            if hasattr(e, 'response') and e.response.status_code == 500:
                logger.warning(f"500 Server Error for {url}, attempt {attempt + 1}/{max_retries}")
                if attempt < max_retries - 1:
                    wait_time = 30 * (attempt + 1)
                    logger.info(f"Waiting {wait_time} seconds before retry with new proxy...")
                    request_data = create_scrape_request(config, url, attempt + 1)
                    time.sleep(wait_time)
                    continue
                else:
                    adaptive_crawler.record_issue(url, 'error')
                    return {}
            
            # Check for specific error patterns
            if 'balance' in error_text.lower():
                logger.error(f"API balance issue: {error_text}")
                logger.error("⚠️  CRITICAL: Check your Scrappey API balance/credits")
                return {}
            elif 'rate limit' in error_text.lower():
                logger.warning(f"Rate limit hit, waiting before retry...")
                if attempt < max_retries - 1:
                    time.sleep(60)  # Longer wait for rate limit
                    request_data = create_scrape_request(config, url, attempt + 1)
                    continue
            elif 'invalid api key' in error_text.lower():
                logger.error(f"Invalid API key: {error_text}")
                logger.error("⚠️  CRITICAL: Check your Scrappey API key")
                return {}
            
            logger.error(f"API call failed: {e}")
            
            if attempt < max_retries - 1:
                request_data = create_scrape_request(config, url, attempt + 1)
                time.sleep(15)
                continue
            return {}
            
        except requests.exceptions.ConnectionError as e:
            logger.error(f"Connection error for {url}: {e}")
            proxy_rotator.record_failure(domain, current_proxy)
            
            if attempt < max_retries - 1:
                request_data = create_scrape_request(config, url, attempt + 1)
                time.sleep(15)
                continue
            return {}
            
        except Exception as e:
            logger.error(f"Unexpected error for {url}: {str(e)}")
            proxy_rotator.record_failure(domain, current_proxy)
            
            if attempt < max_retries - 1:
                request_data = create_scrape_request(config, url, attempt + 1)
                time.sleep(10)
                continue
            return {}
        
        finally:
            # Explicitly close response to free file descriptor
            if response is not None:
                try:
                    response.close()
                except:
                    pass
    
    logger.error(f"All retry attempts exhausted for {url}")
    return {}

def extract_links_from_text(text: str, base_url: str) -> Set[str]:
    """Extract all links from the scraped text"""
    links = set()
    
    url_patterns = [
        r'href=["\']([^"\']+)["\']',
        r'https?://[^\s<>"{}|\\^`\[\]]+',
        r'/[a-zA-Z0-9-_/]+(?:\.[a-zA-Z]+)?'
    ]
    
    base_domain = urlparse(base_url).netloc
    
    for pattern in url_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if match.startswith('/'):
                full_url = urljoin(base_url, match)
            elif match.startswith('http'):
                full_url = match
            else:
                continue
            
            if urlparse(full_url).netloc == base_domain:
                full_url = full_url.split('#')[0].rstrip('/')
                
                # Skip obvious non-content URLs
                skip_patterns = [
                    r'\.(pdf|jpg|png|gif|zip|doc|docx|xls|xlsx|css|js|json|xml)$',
                    r'/(wp-admin|wp-content|wp-includes|assets|static|dist|build)/',
                    r'/(login|logout|register|signin|signup|cart|checkout)/?$'
                ]
                
                if not any(re.search(pattern, full_url.lower()) for pattern in skip_patterns):
                    links.add(full_url)
    
    return links

def crawl_entire_website_smart(config: CrawlConfig, start_url: str, batch_id: str = None, max_depth: int = 3) -> Dict:
    """Crawl website with depth-based link discovery and GPT filtering at each level"""
    
    all_page_texts = []
    all_discovered_links = set()
    visited_urls = set()
    urls_with_content = {}  # Store URL with its scraped content
    
    # Queue structure: [(url, depth)]
    to_crawl = [(start_url, 1)]
    
    # Track failures per depth
    depth_failures = {1: 0, 2: 0, 3: 0}
    
    logger.info(f"Starting depth-based crawl (max depth: {max_depth})")
    if batch_id:
        batch_logger.log(batch_id, start_url, f"Starting depth-based crawl (max depth: {max_depth})", "INFO")
    
    while to_crawl and len(all_page_texts) < config.max_pages:
        current_url, current_depth = to_crawl.pop(0)
        
        # Skip if already visited
        if current_url in visited_urls:
            continue
        
        # Check if we have too many failures at this depth
        if depth_failures.get(current_depth, 0) > 3:
            logger.warning(f"Too many failures at depth {current_depth}, skipping remaining URLs at this depth")
            to_crawl = [(url, d) for url, d in to_crawl if d != current_depth]
            continue
            
        logger.info(f"Crawling at depth {current_depth}: {current_url}")
        
        # Add progressive delay based on depth
        if current_depth > 1:
            delay = config.delay * current_depth
            logger.debug(f"Waiting {delay} seconds before crawling (depth {current_depth})")
            time.sleep(delay)
        
        # Scrape the current page
        request = create_scrape_request(config, current_url, attempt=0)
        result = execute_api_call(config, request)
        
        solution = result.get('solution', {})
        page_text = solution.get('innerText', '')
        html_response = solution.get('response', '')
        
        if not page_text:
            logger.warning(f"No content found for {current_url}")
            depth_failures[current_depth] = depth_failures.get(current_depth, 0) + 1
            continue
        
        # Store URL with its content (truncated for storage efficiency)
        urls_with_content[current_url] = {
            'depth': current_depth,
            'content_preview': page_text[:5000],  # Store first 5000 chars
            'content_length': len(page_text),
            'scraped_at': datetime.now().isoformat()
        }
        
        # Add to results
        all_page_texts.append({
            'url': current_url,
            'text': page_text,
            'success': True,
            'depth': current_depth
        })
        visited_urls.add(current_url)
        
        # If we haven't reached max depth, discover and process child links
        if current_depth < max_depth:
            logger.info(f"Discovering child links from {current_url} (depth {current_depth})")
            
            # Extract all links from current page
            combined_text = page_text + ' ' + html_response
            discovered_links = extract_links_from_text(combined_text, current_url)
            
            # Filter out already visited URLs
            new_links = discovered_links - visited_urls
            
            if new_links:
                logger.info(f"Found {len(new_links)} new links at depth {current_depth}")
                all_discovered_links.update(new_links)
                
                if batch_id:
                    batch_logger.log(batch_id, start_url, 
                        f"Filtering {len(new_links)} links from {current_url} (depth {current_depth})", "INFO")
                
                filter_result = filter_relevant_links_with_gpt_depth_aware(
                    config, 
                    new_links, 
                    current_url,
                    current_depth,
                    page_text
                )
                
                relevant_links = filter_result["filtered_links"]
                
                logger.info(f"GPT selected {len(relevant_links)} relevant links from depth {current_depth}")
                
                # Add all relevant links to crawl queue with incremented depth
                for link in relevant_links:
                    if link not in visited_urls:
                        to_crawl.append((link, current_depth + 1))
                
                if batch_id:
                    batch_logger.log(batch_id, start_url, 
                        f"Added {len(relevant_links)} links to crawl at depth {current_depth + 1}", "INFO")
        
        # Rate limiting between pages
        time.sleep(config.delay)
    
    # Calculate max depth actually reached
    max_depth_reached = max([p.get('depth', 1) for p in all_page_texts]) if all_page_texts else 1
    
    logger.info(f"Crawl complete: {len(all_page_texts)} pages scraped, max depth reached: {max_depth_reached}")
    if batch_id:
        batch_logger.log(batch_id, start_url, 
            f"Completed crawl: {len(all_page_texts)} pages scraped", "INFO")
    
    return {
        'all_pages': all_page_texts,
        'visited_urls': list(visited_urls),
        'base_url': start_url,
        'urls_with_content': urls_with_content,  # Include content mapping
        'total_pages_crawled': len(all_page_texts),
        'max_depth_reached': max_depth_reached
    }


def process_single_ria(config: CrawlConfig, url: str, index: int, total: int, batch_id: str = None) -> Dict:
    """Process single RIA with depth-based crawling"""
    logger.info(f"[{index+1}/{total}] Starting RIA: {url}")
    if batch_id:
        batch_logger.log(batch_id, url, f"Starting RIA processing [{index+1}/{total}]", "INFO")
    
    try:
        # Single-phase depth-based crawl
        website_data = crawl_entire_website_smart(config, url, batch_id, max_depth=3)
        
        # Check if we got any pages at all
        if not website_data.get('all_pages'):
            logger.warning(f"No pages scraped for {url}")
            if batch_id:
                batch_logger.log(batch_id, url, "Failed to crawl website - no pages found", "ERROR")
            return {
                "success": False,
                "url": url,
                "original_url": url,
                "error": "Failed to crawl website - no pages found",
                "processing_index": index,
                "crawl_depth_stats": {
                    "max_depth": 0,
                    "pages_per_depth": {}
                }
            }
        
        # Extract RIA information with enhanced personnel detection
        ria_info = extract_ria_info_from_all_pages(config, website_data)
        
        result = asdict(ria_info)
        result["success"] = True
        result["url"] = url
        result["original_url"] = url
        result["total_pages_crawled"] = len(website_data.get('visited_urls', []))
        result["urls_crawled_with_content"] = website_data.get('urls_with_content', {})
        result["processing_index"] = index
        result["crawl_depth_stats"] = {
            "max_depth": website_data.get('max_depth_reached', 1),
            "pages_per_depth": {}
        }
        
        # Calculate pages per depth safely
        for page in website_data.get('all_pages', []):
            depth = page.get('depth', 1)
            result["crawl_depth_stats"]["pages_per_depth"][depth] = \
                result["crawl_depth_stats"]["pages_per_depth"].get(depth, 0) + 1
        
        logger.info(f"[{index+1}/{total}] Completed RIA: {url}")
        logger.info(f"Crawled to depth {result['crawl_depth_stats']['max_depth']} with distribution: {result['crawl_depth_stats']['pages_per_depth']}")
        
        # Log if no key personnel found
        if not result.get('key_personnel') or len(result.get('key_personnel', [])) == 0:
            logger.warning(f"⚠️ No key personnel found for {url}")
            logger.info("Consider checking URLs containing: /team, /people, /leadership, /about")
        
        if batch_id:
            batch_logger.log(batch_id, url, f"Successfully completed RIA processing", "INFO")
        
        return result
        
    except Exception as e:
        logger.error(f"[{index+1}/{total}] Failed RIA {url}: {e}")
        if batch_id:
            batch_logger.log(batch_id, url, f"Failed with error: {str(e)}", "ERROR")
        return {
            "success": False,
            "url": url,
            "original_url": url,
            "error": str(e),
            "processing_index": index,
            "crawl_depth_stats": {
                "max_depth": 0,
                "pages_per_depth": {}
            }
        }
    
    
def process_multiple_rias_concurrent(
    excel_file: str,
    config: CrawlConfig,
    max_workers: int = 50,
    output_dir: str = 'ria_outputs',
    batch_size: int = 100,  # Add batch size parameter
    start_from: int = 0  # Add resume capability
) -> Dict:
    """Process multiple RIAs in batches with checkpoint saving"""
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Read all URLs
    all_ria_urls = read_ria_urls_from_excel(excel_file)
    
    # Apply start_from offset if resuming
    all_ria_urls = all_ria_urls[start_from:]
    total_urls = len(all_ria_urls)
    
    logger.info(f"Found {total_urls} RIA URLs to process (starting from index {start_from})")
    
    # Process in batches
    num_batches = (total_urls + batch_size - 1) // batch_size
    all_results = []
    
    # Create checkpoint file
    checkpoint_file = os.path.join(output_dir, 'checkpoint.json')
    
    for batch_num in range(num_batches):
        batch_start = batch_num * batch_size
        batch_end = min(batch_start + batch_size, total_urls)
        batch_urls = all_ria_urls[batch_start:batch_end]
        
        # Generate batch ID for this chunk
        batch_id = f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{batch_num+1}of{num_batches}"
        batch_dir = os.path.join(output_dir, batch_id)
        os.makedirs(batch_dir, exist_ok=True)
        
        logger.info(f"\n{'='*60}")
        logger.info(f"Processing Batch {batch_num+1}/{num_batches}")
        logger.info(f"URLs {start_from + batch_start + 1} to {start_from + batch_end} of {start_from + total_urls}")
        logger.info(f"{'='*60}")
        
        results = []
        failed_urls = []
        
        # Process this batch with reasonable concurrency
        with ThreadPoolExecutor(max_workers=min(max_workers, len(batch_urls))) as executor:
            futures = []
            
            for i, url in enumerate(batch_urls):
                global_index = start_from + batch_start + i
                future = executor.submit(process_single_ria, config, url, global_index, start_from + total_urls, batch_id)
                futures.append((future, url))
                time.sleep(1)  # 1 second delay between submissions
            
            # Collect results
            # Collect results
            for future, url in futures:
                result = None
                try:
                    result = future.result(timeout=120)  # Increased timeout
                    results.append(result)
                    
                    # Save individual result immediately with proper error handling
                    if 'url' in result:
                        clean_url = result['url'].replace('https://', '').replace('http://', '').replace('/', '_')
                        result_file = os.path.join(batch_dir, f'{clean_url}_result.json')
                        
                        # Use try-finally to ensure file is closed
                        file_handle = None
                        try:
                            file_handle = open(result_file, 'w', encoding='utf-8')
                            json.dump(result, file_handle, indent=2)
                        except IOError as e:
                            logger.error(f"Failed to write {result_file}: {e}")
                        except Exception as e:
                            logger.error(f"Unexpected error writing {result_file}: {e}")
                        finally:
                            if file_handle is not None:
                                try:
                                    file_handle.close()
                                except:
                                    pass
                                    
                except Exception as e:
                    logger.error(f"Exception for {url}: {str(e)[:200]}")
                    failed_urls.append(url)
                    result = {
                        "success": False,
                        "url": url,
                        "original_url": url,
                        "error": str(e)[:200]
                    }
                    results.append(result)
        
        # Save batch summary
        batch_summary = {
            'batch_id': batch_id,
            'batch_number': batch_num + 1,
            'total_batches': num_batches,
            'urls_in_batch': len(batch_urls),
            'successful': sum(1 for r in results if r.get('success')),
            'failed': len(failed_urls),
            'timestamp': datetime.now().isoformat()
        }
        
        # Save batch output with explicit file handling
        batch_output_file = os.path.join(batch_dir, 'batch_summary.json')
        file_handle = None
        try:
            file_handle = open(batch_output_file, 'w', encoding='utf-8')
            json.dump(batch_summary, file_handle, indent=2)
        except Exception as e:
            logger.error(f"Failed to write batch summary: {e}")
        finally:
            if file_handle is not None:
                try:
                    file_handle.close()
                except:
                    pass
        
        all_results.extend(results)
        
        # Update checkpoint
        # Update checkpoint with explicit file handling
        checkpoint_data = {
            'last_processed_index': start_from + batch_end,
            'last_batch': batch_num + 1,
            'total_batches': num_batches,
            'timestamp': datetime.now().isoformat()
        }
        
        file_handle = None
        try:
            file_handle = open(checkpoint_file, 'w', encoding='utf-8')
            json.dump(checkpoint_data, file_handle)
        except Exception as e:
            logger.error(f"Failed to write checkpoint: {e}")
        finally:
            if file_handle is not None:
                try:
                    file_handle.close()
                except:
                    pass
        
        logger.info(f"Batch {batch_num+1} complete: {batch_summary['successful']} successful, {batch_summary['failed']} failed")
        logger.info(f"Checkpoint saved. Can resume from index {start_from + batch_end}")
        
        # Pause between batches to avoid overwhelming the API
        if batch_num < num_batches - 1:
            logger.info("Waiting 20 seconds before next batch...")
            time.sleep(120)
    
    # Create overall summary
    overall_summary = {
        'total_processed': len(all_results),
        'total_successful': sum(1 for r in all_results if r.get('success')),
        'total_failed': sum(1 for r in all_results if not r.get('success')),
        'batches_processed': num_batches,
        'timestamp': datetime.now().isoformat()
    }
    
    # Save final summary
    # Save final summary with explicit file handling
    final_summary_file = os.path.join(output_dir, 'final_summary.json')
    file_handle = None
    try:
        file_handle = open(final_summary_file, 'w', encoding='utf-8')
        json.dump(overall_summary, file_handle, indent=2)
    except Exception as e:
        logger.error(f"Failed to write final summary: {e}")
    finally:
        if file_handle is not None:
            try:
                file_handle.close()
            except:
                pass
    
    logger.info(f"\n{'='*60}")
    logger.info(f"All batches complete!")
    logger.info(f"Total: {overall_summary['total_processed']}, Successful: {overall_summary['total_successful']}, Failed: {overall_summary['total_failed']}")
    logger.info(f"{'='*60}")
    
    return {
        'summary': overall_summary,
        'detailed_results': all_results
    }



def filter_relevant_links_with_gpt_depth_aware(config: CrawlConfig, links: Set[str], 
                                               parent_url: str, depth: int, 
                                               parent_content: str = "") -> Dict[str, Any]:
    """Use GPT to filter links with awareness of depth and parent content - with batching for large link sets"""
    
    if not links:
        return {"original_links": [], "filtered_links": [], "dropped_links": []}
    
    link_list = list(links)
    
    # Check if we need to batch (if too many links)
    MAX_LINKS_PER_BATCH = 100  # Process max 100 links at a time
    
    if len(link_list) > MAX_LINKS_PER_BATCH:
        logger.info(f"Large link set ({len(link_list)} links) - processing in batches")
        
        # Process in batches
        all_filtered_links = []
        all_dropped_links = []
        
        for i in range(0, len(link_list), MAX_LINKS_PER_BATCH):
            batch = link_list[i:i+MAX_LINKS_PER_BATCH]
            logger.info(f"Processing batch {i//MAX_LINKS_PER_BATCH + 1} ({len(batch)} links)")
            
            # Process this batch
            batch_result = _process_link_batch(config, batch, parent_url, depth, parent_content)
            all_filtered_links.extend(batch_result["filtered_links"])
            all_dropped_links.extend(batch_result["dropped_links"])
        
        return {
            "original_links": link_list,
            "filtered_links": all_filtered_links,
            "dropped_links": all_dropped_links,
            "filter_reason": f"GPT batch filtering at depth {depth}"
        }
    else:
        # Process normally for small link sets
        return _process_link_batch(config, link_list, parent_url, depth, parent_content)

def _process_link_batch(config: CrawlConfig, link_list: List[str], 
                        parent_url: str, depth: int, 
                        parent_content: str = "") -> Dict[str, Any]:
    """Process a single batch of links with GPT"""
    
    client = AzureOpenAI(
        api_key=config.azure_openai_api_key,
        api_version=config.azure_openai_api_version,
        azure_endpoint=config.azure_openai_endpoint
    )
    
    # Limit parent content to avoid token issues
    truncated_parent_content = parent_content[:400000] if parent_content else ""
    
    links_text = "\n".join([f"{i+1}. {link}" for i, link in enumerate(link_list)])
    
    # Context-aware prompt based on depth
    depth_context = {
        1: "homepage links - focus on main navigation pages like about, team, services, contact",
        2: "second-level pages - look for individual team members, specific services, detailed information",
        3: "third-level pages - look for detailed bios, specific expertise, additional contact information"
    }
    
    prompt = f"""
    You are analyzing URLs from level {depth} of an RIA website.
    Parent page: {parent_url}
    Context: These are {depth_context.get(depth, 'child pages')}
    
    From the following list of URLs, identify those likely to contain:
    - Company information and services (any )
    - Team/personnel pages (any  depth)  
    - Individual team member pages (any depth)
    - Detailed bios or profiles (any)
    - Contact page (any depth)
    
    EXCLUDE:
    - Technical pages (css, js, wp-admin, plugins)
    - Media files (images, PDFs unless they're bios)
    - Duplicate or login pages
    - URLs with fragments (#) or query parameters (?)
    
    
    Return the numbers of relevant URLs in JSON:
    {{"relevant_urls": [1, 2, 3, ...]}}
    
    URLs:
    {links_text}
    """
    
    try:
        response = client.chat.completions.create(
            model=config.azure_openai_deployment,
            messages=[
                {"role": "system", "content": "You are an expert at identifying relevant pages on financial websites based on crawl depth."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0
        )
        
        result = json.loads(response.choices[0].message.content)
        relevant_indices = result.get('relevant_urls', [])
        
        relevant_urls = [link_list[i-1] for i in relevant_indices if 0 < i <= len(link_list)]
        dropped_urls = list(set(link_list) - set(relevant_urls))
        
        logger.info(f"Depth {depth}: GPT filtered {len(relevant_urls)} relevant URLs from {len(link_list)} total")
        
        return {
            "original_links": link_list,
            "filtered_links": relevant_urls,
            "dropped_links": dropped_urls,
            "filter_reason": f"GPT filtering at depth {depth}"
        }
        
    except Exception as e:
        logger.error(f"GPT filtering failed at depth {depth}: {e}")




def extract_ria_info_from_all_pages(config: CrawlConfig, website_data: Dict) -> RIAInfo:
    """Enhanced extraction with better personnel detection"""
    
    all_pages = website_data['all_pages']
    base_url = website_data['base_url']
    
    if not all_pages:
        logger.warning(f"No pages scraped from {base_url}")
        return RIAInfo(website_url=base_url, scraped_at=datetime.now().isoformat())
    
    # Prioritize pages likely to contain personnel info
    personnel_keywords = ['team', 'people', 'leadership', 'staff', 'management', 'partners', 'directors', 'executives']
    
    # Sort pages to prioritize those with personnel keywords
    sorted_pages = sorted(all_pages, key=lambda p: 
                         sum(1 for kw in personnel_keywords if kw in p['url'].lower()), 
                         reverse=True)
    
    # Build combined text with personnel pages first
    # Build combined text with personnel pages first - NO TRUNCATION
    combined_text = f"Website: {base_url}\n\n"
    
    for i, page in enumerate(sorted_pages):
        combined_text += f"\n=== PAGE {i+1}: {page['url']} (Depth: {page.get('depth', 1)}) ===\n"
        
        # Include FULL page text (no truncation per page)
        combined_text += page['text'] + "\n"
        
        # Stop if we exceed 80K total characters
        if len(combined_text) > 300000:
            logger.info(f"Reached 80K character limit after {i+1} pages")
            break
    
    client = AzureOpenAI(
        api_key=config.azure_openai_api_key,
        api_version=config.azure_openai_api_version,
        azure_endpoint=config.azure_openai_endpoint
    )
    
    prompt = """
    Analyze this RIA (Registered Investment Advisor) website content and extract comprehensive information.
    
    IMPORTANT: Look carefully for key personnel information which may appear as:
    - Team member names with titles
    - Leadership bios
    - Investment committee members
    - Portfolio managers
    - Partners or principals
    - Advisory board members
    
    Extract:
    1. Company Name
    2. All services offered
    3. Target clients (Family Offices, Multi-Family Offices, UHNW, HNW, etc.)
    4. Investment Strategy and Philosophy
    5. Products/instruments used
    6. ALL key personnel - BE THOROUGH, include everyone mentioned with a professional role
    7. Industry insights or thought leadership themes
    8. Complete contact information
    
    Return as JSON:
    {
        "company_name": "",
        "services": [],
        "target_clients": [],
        "investment_strategy": "",
        "products_instruments": [],
        "key_personnel": [
            {"name": "", "title": "", "bio": "", "contact": ""}
        ],
        "industry_insights": "",
        "contact_info": {"address": "", "phone": "", "email": "", "offices": []}
    }
    
    Content:
    """
    
    try:
        response = client.chat.completions.create(
            model=config.azure_openai_deployment,
            messages=[
                {"role": "system", "content": "You are an expert at extracting comprehensive RIA firm information. Pay special attention to identifying ALL personnel mentioned."},
                {"role": "user", "content": prompt + combined_text}
            ],
            response_format={"type": "json_object"},
            temperature=0
        )
        
        extracted_data = json.loads(response.choices[0].message.content)
        
        # Log extraction summary
        logger.info(f"Extraction complete for {base_url}:")
        logger.info(f"  - Company: {extracted_data.get('company_name', 'Unknown')}")
        logger.info(f"  - Services: {len(extracted_data.get('services', []))} found")
        logger.info(f"  - Personnel: {len(extracted_data.get('key_personnel', []))} found")
        
        return RIAInfo(
            company_name=extracted_data.get('company_name', ''),
            services=extracted_data.get('services', []),
            target_clients=extracted_data.get('target_clients', []),
            investment_strategy=extracted_data.get('investment_strategy', ''),
            products_instruments=extracted_data.get('products_instruments', []),
            key_personnel=extracted_data.get('key_personnel', []),
            industry_insights=extracted_data.get('industry_insights', ''),
            contact_info=extracted_data.get('contact_info', {}),
            website_url=base_url,
            scraped_at=datetime.now().isoformat(),
            pages_analyzed=len(all_pages),
            all_page_data=[{'url': p['url'], 'depth': p.get('depth', 1)} for p in all_pages]
        )
        
    except Exception as e:
        logger.error(f"GPT extraction failed: {e}")
        return RIAInfo(website_url=base_url, scraped_at=datetime.now().isoformat())

def test_scrappey_api(config: CrawlConfig):
    """Test if Scrappey API is working with a simple site"""
    
    print(f"\n{'='*60}")
    print("Testing Scrappey API")
    print(f"{'='*60}")
    
    # Test with a simple, known-to-work site
    test_urls = [
        "https://google.com",  # Simple test site
        "https://www.franklintempleton.com",  # From your working example
    ]
    
    for test_url in test_urls:
        print(f"\nTesting: {test_url}")
        
        request_data = {
            'headers': {'Content-Type': 'application/json'},
            'params': {'key': config.api_key},
            'json_data': {
                'cmd': 'request.get',
                'url': test_url
            }
        }
        
        try:
            response = requests.post(
                config.base_url,
                params=request_data['params'],
                headers=request_data['headers'],
                json=request_data['json_data'],
                timeout=40
            )
            
            if response.status_code == 200:
                result = response.json()
                solution = result.get('solution', {})
                inner_text = solution.get('innerText', '')
                
                if inner_text:
                    print(f"  ✅ Success! Got {len(inner_text)} characters")
                    print(f"  First 100 chars: {inner_text[:100]}...")
                else:
                    print(f"  ⚠️ Response received but no content")
            else:
                print(f"  ❌ HTTP {response.status_code}: {response.text[:200]}")
                
        except Exception as e:
            print(f"  ❌ Error: {str(e)}")
    
    # Check API balance/status
    print(f"\n{'='*60}")
    print("API Key Info:")
    print(f"  Key: {config.api_key[:10]}...{config.api_key[-5:]}")
    print(f"  Endpoint: {config.base_url}")
    print(f"{'='*60}")


def print_proxy_statistics():
    """Print statistics about proxy usage"""
    logger.info("\n" + "="*60)
    logger.info("PROXY USAGE STATISTICS")
    logger.info("="*60)
    
    with proxy_rotator.lock:
        for domain, proxy_stats in proxy_rotator.domain_proxy_success.items():
            logger.info(f"\nDomain: {domain}")
            for proxy, success_count in proxy_stats.items():
                failures = proxy_rotator.domain_proxy_failures[domain][proxy]
                total = success_count + failures
                success_rate = (success_count / total * 100) if total > 0 else 0
                logger.info(f"  {proxy}: {success_count} success, {failures} failures ({success_rate:.1f}% success rate)")
    
    logger.info("="*60)
    
    
# Usage
if __name__ == "__main__":
    config = CrawlConfig(
        api_key='sbV103de9UccEp2YUCqBspkUn1aeHR5TOJdsFguS9gq30xlBqpECgaNVy4aw',
        azure_openai_api_key=os.getenv('AZURE_OPENAI_API_KEY'),
        azure_openai_endpoint=os.getenv('AZURE_OPENAI_ENDPOINT'),
        azure_openai_api_version=os.getenv('AZURE_OPENAI_API_VERSION'),
        azure_openai_deployment=os.getenv('AZURE_OPENAI_DEPLOYMENT_NAME'),
        max_pages=20,
        delay=10.0
    )
    
    mode = 'batch'  # 'single', 'batch', or 'test'
    
    if mode == 'test':
        # TEST MODE - Check if API is working
        test_scrappey_api(config)   
        
    elif mode == 'single':
        # SINGLE URL MODE
        single_url = "https://www.franklintempleton.com"
        
        # Clean the URL
        single_url = clean_single_url(single_url)
        
        print(f"\n{'='*60}")
        print(f"Processing Single URL: {single_url}")
        print(f"{'='*60}")
        
        # Create batch ID for logging
        batch_id = f"single_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Process single URL
        result = process_single_ria(config, single_url, 0, 1, batch_id)
        
        # Save result
        output_dir = 'ria_outputs'
        os.makedirs(output_dir, exist_ok=True)
        
        clean_filename = single_url.replace('https://', '').replace('http://', '').replace('/', '_').replace(':', '')
        output_file = os.path.join(output_dir, f'{clean_filename}_result.json')
        
        with open(output_file, 'w') as f:
            json.dump(result, f, indent=2)
        
        # Print summary
        print(f"\n{'='*60}")
        print(f"Processing Complete!")
        print(f"{'='*60}")
        print(f"URL: {single_url}")
        print(f"Success: {result.get('success', False)}")
        print(f"Pages crawled: {result.get('total_pages_crawled', 0)}")
        
        if result.get('success'):
            print(f"\nExtracted Information:")
            print(f"Company: {result.get('company_name', 'N/A')}")
            print(f"Services: {len(result.get('services', [])) if result.get('services') else 0} services found")
            print(f"Key Personnel: {len(result.get('key_personnel', [])) if result.get('key_personnel') else 0} people found")
            print(f"Target Clients: {', '.join(result.get('target_clients', [])) if result.get('target_clients') else 'N/A'}")
        
        if result.get('error'):
            print(f"Error: {result.get('error')}")
        
        print(f"\nResults saved to: {output_file}")
        print(f"{'='*60}")
        
    elif mode == 'batch':
        # BATCH MODE FROM EXCEL
        excel_file = '/Users/nikitakumar/Documents/RIA_Web_Scraping/RIA_websites_company_websites_only.xlsx'
        
        # Check if we're resuming from checkpoint
        checkpoint_file = 'ria_outputs/checkpoint.json'
        start_from = 11875
        
        if os.path.exists(checkpoint_file):
            with open(checkpoint_file, 'r') as f:
                checkpoint = json.load(f)
                resume_choice = input(f"Resume from index {checkpoint['last_processed_index']}? (y/n): ")
                if resume_choice.lower() == 'y':
                    start_from = checkpoint['last_processed_index']
        
        print(f"{'='*60}")
        print(f"Processing Batch from Excel")
        print(f"File: {excel_file}")
        print(f"Starting from index: {start_from}")
        print(f"{'='*60}")
        
        results = process_multiple_rias_concurrent(
            excel_file=excel_file,
            config=config,
            max_workers=150,  # CHANGED from 190 to 5
            batch_size=190,  # Process 100 at a time
            start_from=start_from  # Resume capability
        )
        
        # Print batch summary
        print(f"\n{'='*60}")
        print(f"Batch Processing Complete!")
        print(f"{'='*60}")
        print(f"Batch ID: {results.get('batch_id', 'N/A')}")
        print(f"Total RIAs: {results['summary']['total_processed']}")
        print(f"Successful: {results['summary']['successful']}")
        print(f"Failed: {results['summary']['failed']}")
        
        # Show details for each URL
        if results['summary'].get('ria_status'):
            print(f"\nDetailed Status:")
            for url, status in results['summary']['ria_status'].items():
                status_symbol = "✓" if status['success'] else "✗"
                pages = status['pages_crawled']
                print(f"  {status_symbol} {url}: {pages} pages")
                if status.get('error'):
                    print(f"    Error: {status['error']}")
        
        # Show extraction summary
        print(f"\nExtraction Summary:")
        successful_extractions = [r for r in results.get('detailed_results', []) if r.get('success')]
        for result in successful_extractions[:5]:  # Show first 5
            print(f"\n  {result.get('company_name', result.get('url', 'Unknown'))}:")
            print(f"    - Services: {len(result.get('services', []))} found")
            print(f"    - Personnel: {len(result.get('key_personnel', []))} found")
            print(f"    - Pages: {result.get('total_pages_crawled', 0)} crawled")
        
        if len(successful_extractions) > 5:
            print(f"\n  ... and {len(successful_extractions) - 5} more")
        
        print(f"\nResults saved to: ria_outputs/{results.get('batch_id', 'batch_output')}/")
        print(f"{'='*60}")
    
    else:
        print("Invalid mode. Please set mode to 'single', 'batch', or 'test'")
        
        
    print_proxy_statistics()
    
    # Close all sessions
    session_pool.close_all_sessions()
    logger.info("All sessions closed. Exiting.")