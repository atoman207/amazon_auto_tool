#!/usr/bin/env python3
"""
Amazon Business Category Search Tool
Searches for discounted products by category keywords and displays results
"""

import os
import re
import base64
import time
import sys
import json
from pathlib import Path
from datetime import datetime, timedelta

# Check and import required packages
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
except ImportError as e:
    print("\n" + "="*70)
    print("[ERROR] Playwright is not installed!")
    print("="*70)
    print("Please install it by running:")
    print("  pip install playwright")
    print("  playwright install chrome")
    print("="*70 + "\n")
    sys.exit(1)

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError as e:
    print("\n" + "="*70)
    print("[ERROR] Gmail API packages are not installed!")
    print("="*70)
    print("Please install them by running:")
    print("  pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    print("="*70)
    print(f"Detailed error: {e}")
    print("="*70 + "\n")
    sys.exit(1)

try:
    from dotenv import load_dotenv
except ImportError as e:
    print("\n" + "="*70)
    print("[ERROR] python-dotenv is not installed!")
    print("="*70)
    print("Please install it by running:")
    print("  pip install python-dotenv")
    print("="*70)
    print(f"Detailed error: {e}")
    print("="*70 + "\n")
    sys.exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Load environment variables
load_dotenv()

# Amazon credentials
AMAZON_EMAIL = os.getenv('AMAZON_EMAIL', 'gocean0807@gmail.com')
AMAZON_PASSWORD = os.getenv('AMAZON_PASSWORD')

if not AMAZON_PASSWORD:
    print("\n" + "="*70)
    print("[ERROR] AMAZON_PASSWORD not found in .env file")
    print("="*70)
    print("Please create a .env file in the project root with:")
    print("  AMAZON_EMAIL=your_email@gmail.com")
    print("  AMAZON_PASSWORD=your_password")
    print("="*70 + "\n")
    sys.exit(1)

# Gmail API settings
GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
GMAIL_CREDENTIALS_FILE = Path('data/client_secret_446842116198-amdijg8d7tb7rff25o4514r19pp1d8o9.apps.googleusercontent.com.json')
GMAIL_TOKEN_FILE = Path('token.json')

# Session file
SESSION_FILE = "amazon_session.json"

# Amazon URLs (Japanese site)
AMAZON_LOGIN_URL = "https://www.amazon.co.jp/ap/signin?openid.pape.max_auth_age=900&openid.return_to=https%3A%2F%2Fwww.amazon.co.jp%2Fgp%2Fyourstore%2Fhome%3Fpath%3D%252Fgp%252Fyourstore%252Fhome%26signIn%3D1%26useRedirectOnSuccess%3D1%26action%3Dsign-out%26ref_%3Dabn_yadd_sign_out&openid.assoc_handle=jpflex&openid.mode=checkid_setup&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
BUSINESS_DISCOUNTS_URL = "https://www.amazon.co.jp/ab/business-discounts?ref_=abn_cs_savings_guide&pd_rd_r=242d4956-5f68-4e46-bc0d-0fc896eaadf4&pd_rd_w=jg2kX&pd_rd_wg=rMzwy"

# Search keywords (categories to search)
SEARCH_KEYWORDS = [
    "IT関連機器",           # IT Equipment
    "医療用品・消耗品",     # Medical Supplies
    "日用品・食品・飲料"    # Daily Necessities
]

# Timeouts and delays
TIMEOUT_MS = 30000
DELAY_AFTER_CLICK = 2

# Selectors for Amazon login (Japanese site)
EMAIL_SELECTORS = [
    "xpath=/html/body/div[1]/div[1]/div[2]/div/div/div/div/span/form/div[1]/input",
    "#ap_email", 
    "input[name='email']", 
    "input[type='email']", 
    "input#ap_email"
]
CONTINUE_SELECTORS = [
    "input[aria-labelledby='continue-announce']",
    "#continue", 
    "input#continue", 
    "input[type='submit']"
]
PASSWORD_SELECTORS = [
    "xpath=/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div[2]/div/form/div/div[1]/input",
    "#ap_password", 
    "input[name='password']", 
    "input[type='password']", 
    "input#ap_password"
]
SIGNIN_SELECTORS = [
    "xpath=/html/body/div[1]/div[1]/div[2]/div/div[2]/div/div[2]/div/form/div/div[2]/span/span/input",
    "#signInSubmit", 
    "input#signInSubmit", 
    "input[type='submit']", 
    "button[type='submit']"
]
OTP_SELECTORS = [
    "xpath=/html/body/div[1]/div[2]/div/div/div[3]/form/div[1]/div/div/span[1]/div/input",
    "#auth-mfa-otpcode", 
    "input[name='otpCode']", 
    "input[autocomplete='one-time-code']",
    "input[name='code']",
    "input[type='tel']",
    "input#auth-mfa-otpcode",
    "input[inputmode='numeric']"
]
OTP_SUBMIT_SELECTORS = [
    "xpath=/html/body/div[1]/div[2]/div/div/div[3]/form/div[7]/span/span/input",
    "#auth-signin-button", 
    "input[type='submit']", 
    "button[type='submit']", 
    "input#auth-signin-button"
]

# Selectors for search functionality
SEARCH_INPUT = "xpath=/html/body/div[1]/header/div/div[1]/div[2]/div[1]/form/div[2]/div[1]/input"
SEARCH_BUTTON = "xpath=/html/body/div[1]/header/div/div[1]/div[2]/div[1]/form/div[3]/div/span/input"

# Pagination selector
NEXT_PAGE_BUTTON = ".s-pagination-next, a.s-pagination-item.s-pagination-next, .a-pagination .a-last a"


# ============================================================================
# GMAIL API FUNCTIONS
# ============================================================================

def get_gmail_service():
    """
    Authenticate and return Gmail API service
    
    Returns:
        Gmail API service object
    """
    creds = None
    
    # Load existing token if available
    if GMAIL_TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(GMAIL_TOKEN_FILE), GMAIL_SCOPES)
        except Exception as e:
            print(f"[WARNING] Could not load token: {e}")
            creds = None
    
    # If no valid credentials, do OAuth flow (one-time)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("[INFO] Refreshing expired token...")
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"[WARNING] Token refresh failed: {e}")
                creds = None
        
        if not creds:
            if not GMAIL_CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"\n[ERROR] Gmail credentials file not found: {GMAIL_CREDENTIALS_FILE}\n"
                    "Please download OAuth credentials from Google Cloud Console\n"
                    "and place it in the 'data' folder."
                )
            
            print("\n" + "="*60)
            print("GMAIL API AUTHORIZATION REQUIRED")
            print("="*60)
            print("A browser window will open for Google authorization.")
            print("This is required only the first time (token.json is saved).")
            print("="*60 + "\n")
            
            flow = InstalledAppFlow.from_client_secrets_file(
                str(GMAIL_CREDENTIALS_FILE), GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next time
        with open(GMAIL_TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        print(f"[SUCCESS] Gmail authentication saved to {GMAIL_TOKEN_FILE}")
    
    return build('gmail', 'v1', credentials=creds)


def extract_otp_from_html(html_text):
    """
    Extract 6-digit OTP code from HTML email
    
    Args:
        html_text: Email HTML body
    
    Returns:
        6-digit OTP code as string, or None if not found
    """
    if not html_text:
        return None
    
    try:
        # Method 1: Try to find the specific table structure
        table_pattern = r'<table[^>]*>.*?<tbody[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<td[^>]*>.*?<div[^>]*>.*?<span[^>]*>(\d{6})</span>'
        match = re.search(table_pattern, html_text, re.DOTALL | re.IGNORECASE)
        if match:
            otp = match.group(1)
            if len(otp) == 6 and otp.isdigit():
                return otp
        
        # Method 2: Find all spans with 6-digit numbers
        span_pattern = r'<span[^>]*>(\d{6})</span>'
        spans = re.findall(span_pattern, html_text, re.IGNORECASE)
        for span_text in spans:
            if len(span_text) == 6 and span_text.isdigit():
                span_index = html_text.find(f'<span>{span_text}</span>')
                if span_index > 0:
                    context = html_text[max(0, span_index-500):span_index+100]
                    if 'table' in context.lower() and 'tbody' in context.lower():
                        return span_text
        
    except Exception as e:
        print(f"    [WARNING] HTML parsing error: {e}")
    
    return None


def extract_otp_from_text(text):
    """
    Extract 6-digit OTP code from email text
    
    Args:
        text: Email body text
    
    Returns:
        6-digit OTP code as string, or None if not found
    """
    if not text:
        return None
    
    # First try HTML extraction
    if '<html' in text.lower() or '<body' in text.lower() or '<table' in text.lower():
        html_otp = extract_otp_from_html(text)
        if html_otp:
            return html_otp
    
    # Then try regex patterns for plain text
    patterns = [
        r'確認コード(?:は|:|：)(?:次のとおりです)?(?:\s*[:：]\s*)?(\d{6})',
        r'verification\s+code(?:\s+is)?(?:\s*[:：]\s*)?(\d{6})',
        r'コード(?:\s*[:：]\s*)(\d{6})',
        r'(?:^|\s)(\d{6})(?:\s|$)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            otp = match.group(1)
            if len(otp) == 6 and otp.isdigit():
                return otp
    
    return None


def decode_email_body(payload):
    """
    Decode email body from Gmail API message payload
    
    Args:
        payload: Message payload from Gmail API
    
    Returns:
        Decoded email body text
    """
    html_body = ""
    text_body = ""
    
    if 'body' in payload and 'data' in payload['body']:
        body_data = payload['body']['data']
        decoded = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
        if '<html' in decoded.lower() or '<body' in decoded.lower() or '<table' in decoded.lower():
            html_body = decoded
        else:
            text_body = decoded
    
    if 'parts' in payload:
        for part in payload['parts']:
            mime_type = part.get('mimeType', '')
            
            if 'parts' in part:
                nested_result = decode_email_body(part)
                if '<html' in nested_result.lower() or '<table' in nested_result.lower():
                    html_body += nested_result
                else:
                    text_body += nested_result
            
            elif mime_type == 'text/html':
                if 'data' in part['body']:
                    part_data = part['body']['data']
                    decoded = base64.urlsafe_b64decode(part_data).decode('utf-8', errors='ignore')
                    html_body += decoded + "\n"
            
            elif mime_type == 'text/plain':
                if 'data' in part['body']:
                    part_data = part['body']['data']
                    decoded = base64.urlsafe_b64decode(part_data).decode('utf-8', errors='ignore')
                    text_body += decoded + "\n"
    
    return html_body if html_body else text_body


def get_amazon_otp_from_gmail(max_age_minutes=5, max_retries=12, retry_delay=5):
    """
    Get latest Amazon OTP code from Gmail
    
    Args:
        max_age_minutes: Only check emails from last N minutes
        max_retries: Maximum number of retry attempts
        retry_delay: Seconds to wait between retries
    
    Returns:
        6-digit OTP code as string, or None if not found
    """
    
    print("\n" + "="*60)
    print("RETRIEVING AMAZON OTP FROM GMAIL")
    print("="*60)
    
    try:
        service = get_gmail_service()
        print("[SUCCESS] Connected to Gmail API")
    except Exception as e:
        print(f"[ERROR] Failed to connect to Gmail: {e}")
        return None
    
    since_time = int((datetime.now() - timedelta(minutes=max_age_minutes)).timestamp())
    
    query = (
        f'(from:amazon.co.jp OR from:account-update@amazon.co.jp OR from:no-reply@amazon.co.jp '
        f'OR from:auto-confirm@amazon.co.jp) after:{since_time}'
    )
    
    print(f"\n[INFO] Searching for Amazon verification emails from last {max_age_minutes} minutes...")
    
    for attempt in range(1, max_retries + 1):
        try:
            results = service.users().messages().list(
                userId='me',
                q=query,
                maxResults=10
            ).execute()
            
            messages = results.get('messages', [])
            
            if not messages:
                if attempt < max_retries:
                    print(f"[Attempt {attempt}/{max_retries}] No email found yet, waiting {retry_delay}s...")
                    time.sleep(retry_delay)
                    continue
                else:
                    print(f"\n[WARNING] No Amazon email found after {max_retries} attempts")
                    return None
            
            print(f"[SUCCESS] Found {len(messages)} email(s) from Amazon")
            
            for idx, message in enumerate(messages, 1):
                try:
                    msg = service.users().messages().get(
                        userId='me',
                        id=message['id'],
                        format='full'
                    ).execute()
                    
                    headers = msg['payload'].get('headers', [])
                    subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
                    from_email = next((h['value'] for h in headers if h['name'].lower() == 'from'), 'Unknown')
                    
                    print(f"\n  Checking Email {idx}:")
                    print(f"    From: {from_email[:50]}")
                    print(f"    Subject: {subject[:60]}...")
                    
                    try:
                        internal_ms = int(msg.get("internalDate", "0"))
                        if internal_ms:
                            age_seconds = (time.time() - (internal_ms / 1000.0))
                            if age_seconds > (max_age_minutes * 60):
                                print("    [INFO] Skipping (too old)")
                                continue
                    except Exception:
                        pass

                    body_text = decode_email_body(msg['payload'])
                    otp = extract_otp_from_text(body_text)
                    
                    if otp:
                        print(f"    [SUCCESS] Found OTP: {otp}")
                        print("\n" + "="*60)
                        return otp
                    else:
                        print(f"    [INFO] No OTP in this email")
                
                except HttpError as e:
                    print(f"    [ERROR] Failed to read email: {e}")
                    continue
            
            if attempt < max_retries:
                print(f"\n[Attempt {attempt}/{max_retries}] OTP not found, waiting {retry_delay}s...")
                time.sleep(retry_delay)
            
        except HttpError as e:
            print(f"[ERROR] Gmail API error: {e}")
            if attempt < max_retries:
                time.sleep(retry_delay)
            continue
    
    print(f"\n[WARNING] Could not find OTP after {max_retries} attempts")
    print("="*60)
    return None


# ============================================================================
# BROWSER AUTOMATION FUNCTIONS
# ============================================================================

def find_first_visible(page, selectors, timeout=5000):
    """Find the first visible element from a list of selectors"""
    for selector in selectors:
        try:
            locator = page.locator(selector)
            if locator.count() > 0:
                element = locator.first
                if element.is_visible(timeout=timeout):
                    return element
        except Exception:
            continue
    return None


def human_click(locator, delay_after=0.3):
    """
    Click with slow, visible, human-like mouse movement
    """
    try:
        locator.scroll_into_view_if_needed(timeout=TIMEOUT_MS)
        time.sleep(0.3)  # Pause after scrolling into view
    except Exception:
        pass
    
    try:
        box = locator.bounding_box()
        if box:
            # Calculate center position
            x = box["x"] + box["width"] / 2
            y = box["y"] + box["height"] / 2
            
            # Move mouse slowly and visibly
            locator.page.mouse.move(x, y)
            time.sleep(0.5)  # Increased from 0.1 to 0.5 for visibility
    except Exception:
        pass
    
    # Brief pause before clicking
    time.sleep(0.3)
    locator.click()
    time.sleep(delay_after)


def wait_for_page_load(page, timeout=TIMEOUT_MS):
    """Wait for page to load"""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout)
    except PWTimeoutError:
        print("[WARNING] Page load timeout, continuing...")


def scroll_product_page_slowly(page, scroll_times=20, scroll_delay=2.0):
    """
    Slowly and smoothly scroll down the products page to display all products
    
    Args:
        page: Playwright page object
        scroll_times: Number of times to scroll (default: 20 for smoother scrolling)
        scroll_delay: Delay between scrolls in seconds (default: 2.0 for visible scrolling)
    """
    print(f"\n[INFO] ゆっくりスクロール開始 ({scroll_times}回スクロール、各{scroll_delay}秒間隔)...")
    try:
        for i in range(scroll_times):
            # Smaller scroll amount for smoother, more visible scrolling
            page.mouse.wheel(0, 300)  # Reduced from 500 to 300 pixels
            time.sleep(scroll_delay)
            
            # Print progress more frequently
            if (i + 1) % 3 == 0:
                print(f"[INFO] スクロール進捗: {i + 1}/{scroll_times} 回")
        
        print("[SUCCESS] スクロール完了")
        time.sleep(1)  # Brief pause after scrolling completes
    except Exception as e:
        print(f"[WARNING] スクロール中にエラー: {e}")


def check_and_navigate_next_page(page):
    """
    Check if there's a next page and navigate to it with visible, slow actions
    
    Args:
        page: Playwright page object
    
    Returns:
        True if navigated to next page, False if no next page
    """
    try:
        print("\n[INFO] 次のページを探しています...")
        time.sleep(1)  # Pause before searching
        
        # Try multiple selectors for next page button
        next_selectors = [
            ".s-pagination-next:not(.s-pagination-disabled)",
            "a.s-pagination-item.s-pagination-next:not(.s-pagination-disabled)",
            ".a-pagination .a-last:not(.a-disabled) a",
            "li.a-last:not(.a-disabled) a"
        ]
        
        for selector in next_selectors:
            try:
                next_button = page.locator(selector).first
                if next_button.count() > 0 and next_button.is_visible(timeout=2000):
                    print(f"[SUCCESS] 次のページボタンを発見: {selector}")
                    
                    # Scroll to button to make it visible
                    try:
                        next_button.scroll_into_view_if_needed(timeout=3000)
                        print("[INFO] ボタンまでスクロール完了")
                        time.sleep(1)
                    except Exception:
                        pass
                    
                    # Highlight the button by hovering over it
                    try:
                        box = next_button.bounding_box()
                        if box:
                            x = box["x"] + box["width"] / 2
                            y = box["y"] + box["height"] / 2
                            page.mouse.move(x, y)
                            print("[INFO] 次のページボタンにマウスホバー中...")
                            time.sleep(1.5)  # Hover for visibility
                    except Exception:
                        pass
                    
                    # Click the button slowly
                    print("[INFO] 次のページボタンをクリックします...")
                    time.sleep(0.5)  # Brief pause before clicking
                    human_click(next_button, delay_after=3.0)
                    
                    # Wait longer for page to load
                    print("[INFO] 次のページの読み込み待機中...")
                    time.sleep(4)  # Increased from 3 to 4 seconds
                    wait_for_page_load(page)
                    time.sleep(1)  # Additional pause after page load
                    
                    print("[SUCCESS] 次のページへ移動完了")
                    return True
            except Exception as e:
                print(f"[DEBUG] セレクタ {selector} で失敗: {e}")
                continue
        
        print("[INFO] 次のページが見つかりません（ページネーション終了）")
        return False
        
    except Exception as e:
        print(f"[INFO] ページネーション検索中にエラー: {e}")
        return False


def search_and_display_products(page, keyword):
    """
    Search for a keyword and display all products with pagination
    
    Args:
        page: Playwright page object
        keyword: Search keyword
    """
    print("\n" + "="*60)
    print(f"SEARCHING FOR: {keyword}")
    print("="*60)
    
    try:
        # Find search input
        print("\n[1/3] Entering search keyword...")
        search_input = page.locator(SEARCH_INPUT).first
        
        if search_input.count() == 0:
            print("[ERROR] Search input field not found")
            return False
        
        # Clear existing text and enter keyword
        search_input.clear()
        search_input.fill(keyword)
        time.sleep(0.5)
        print(f"[SUCCESS] Entered keyword: {keyword}")
        
        # Click search button
        print("\n[2/3] Clicking search button...")
        search_button = page.locator(SEARCH_BUTTON).first
        
        if search_button.count() == 0:
            print("[ERROR] Search button not found")
            return False
        
        human_click(search_button, delay_after=2.0)
        wait_for_page_load(page)
        time.sleep(2)
        print("[SUCCESS] Search executed")
        
        # Display all products with pagination
        print("\n[3/3] 全商品を表示します（ゆっくりスクロール）...")
        page_number = 1
        
        while True:
            print(f"\n{'='*60}")
            print(f"ページ {page_number} を表示中")
            print(f"{'='*60}")
            
            # Wait a moment before scrolling
            time.sleep(1)
            
            # Scroll slowly and smoothly through current page
            # Increased scroll_times for more visible scrolling, longer delay for clarity
            scroll_product_page_slowly(page, scroll_times=20, scroll_delay=2.0)
            
            # Pause after scrolling completes
            print(f"\n[INFO] ページ {page_number} のスクロール完了")
            time.sleep(2)
            
            # Check for next page with visible actions
            has_next = check_and_navigate_next_page(page)
            
            if not has_next:
                break
            
            page_number += 1
            
            # Brief pause between pages
            print(f"\n[INFO] 次のページ（{page_number}）を準備中...")
            time.sleep(1.5)
            
            # Safety limit to prevent infinite loops
            if page_number > 20:
                print(f"\n[WARNING] 最大ページ数（20ページ）に到達しました")
                break
        
        print("\n" + "="*60)
        print(f"[SUCCESS] Completed search for: {keyword}")
        print(f"[INFO] Total pages displayed: {page_number}")
        print("="*60)
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Failed to search for '{keyword}': {e}")
        import traceback
        traceback.print_exc()
        return False


def login_to_amazon(page, context):
    """
    Login to Amazon with automatic OTP retrieval from Gmail
    
    Args:
        page: Playwright page object
        context: Playwright browser context
        
    Returns:
        True if login successful, False otherwise
    """
    
    print("\n" + "="*60)
    print("STEP 1: AMAZON LOGIN")
    print("="*60)
    
    try:
        print("\n[1/5] Navigating to Amazon Japan login page...")
        page.goto(AMAZON_LOGIN_URL, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
        time.sleep(2)
        print("[SUCCESS] Loaded Amazon Japan login page")
        
        # Check for Passkey modal and close it if present
        print("\n[INFO] Checking for Passkey modal...")
        try:
            close_btn = page.locator('button:has-text("閉じる"), [aria-label="閉じる"], button:has-text("Close")').first
            if close_btn.is_visible(timeout=3000):
                print("[INFO] Passkey modal detected. Closing...")
                human_click(close_btn)
                time.sleep(1)
                print("[SUCCESS] Closed Passkey modal")
                
                login_btn = page.locator('a:has-text("ログイン"), button:has-text("ログイン"), a:has-text("Login"), button:has-text("Login")').first
                if login_btn.is_visible(timeout=2000):
                    print("[INFO] Clicking Login button after modal close...")
                    human_click(login_btn)
                    time.sleep(2)
            else:
                print("[INFO] No Passkey modal found")
        except Exception as e:
            print(f"[INFO] Passkey check skipped: {e}")
        
        # Enter email
        print("\n[2/5] Entering email...")
        email_input = find_first_visible(page, EMAIL_SELECTORS)
        if not email_input:
            raise RuntimeError("Could not find email input field")
        
        email_input.fill(AMAZON_EMAIL)
        time.sleep(0.5)
        print(f"[SUCCESS] Entered email: {AMAZON_EMAIL}")
        
        # Set up dialog handler
        print("[INFO] Setting up dialog handler for passkey alert...")
        dialog_handled = False
        
        def handle_dialog(dialog):
            nonlocal dialog_handled
            print(f"[INFO] Browser dialog detected: {dialog.type}")
            print(f"[INFO] Dialog message: {dialog.message}")
            if "passkey" in dialog.message.lower() or "パスキー" in dialog.message:
                print("[SUCCESS] Passkey alert detected - dismissing...")
                dialog.dismiss()
                dialog_handled = True
            else:
                print("[WARNING] Unknown dialog - accepting...")
                dialog.accept()
        
        page.on("dialog", handle_dialog)
        
        # Click continue
        continue_btn = find_first_visible(page, CONTINUE_SELECTORS)
        if continue_btn:
            print("[INFO] Clicking continue button...")
            human_click(continue_btn, delay_after=0.5)
            wait_for_page_load(page)
            time.sleep(1)
            
        if dialog_handled:
            print("[SUCCESS] Passkey alert was automatically dismissed")
        else:
            print("[INFO] No passkey alert appeared")
        
        # Wait for password field
        print("\n" + "="*70)
        print("⚠️  PASSKEY MODAL MAY APPEAR - PLEASE CLOSE IT MANUALLY")
        print("="*70)
        print("[ACTION REQUIRED] If a passkey alert appears:")
        print("  1. Click the '閉じる' (Close) button")
        print("  2. The script will continue automatically")
        print("="*70)
        
        print("\n[INFO] Waiting for password field to become accessible...")
        password_accessible = False
        max_wait_time = 120
        check_interval = 1
        elapsed_time = 0
        
        while not password_accessible and elapsed_time < max_wait_time:
            try:
                password_test = find_first_visible(page, PASSWORD_SELECTORS, timeout=1000)
                
                if password_test:
                    try:
                        password_test.focus(timeout=1000)
                        password_test.press_sequentially("", timeout=1000)
                        password_accessible = True
                        print(f"\n[SUCCESS] Password field accessible after {elapsed_time}s!")
                        break
                    except Exception:
                        pass
                
                time.sleep(check_interval)
                elapsed_time += check_interval
                
                if elapsed_time % 5 == 0 and elapsed_time > 0:
                    print(f"[INFO] Still waiting... ({elapsed_time}s elapsed)")
                    
            except Exception as e:
                time.sleep(check_interval)
                elapsed_time += check_interval
        
        if password_accessible:
            print("[SUCCESS] Password field confirmed accessible")
        else:
            print(f"[WARNING] Timeout after {max_wait_time}s - continuing anyway")
        
        try:
            page.remove_listener("dialog", handle_dialog)
        except:
            pass
        
        # Enter password
        print("\n[3/5] Entering password...")
        password_input = None
        for retry in range(5):
            password_input = find_first_visible(page, PASSWORD_SELECTORS, timeout=3000)
            if password_input:
                print(f"[SUCCESS] Password field found (attempt {retry + 1}/5)")
                break
            else:
                if retry < 4:
                    print(f"[INFO] Waiting... (attempt {retry + 1}/5)")
                    time.sleep(1)
                else:
                    raise RuntimeError("Could not find password input field")
        
        if not password_input:
            raise RuntimeError("Could not find password input field")
        
        password_input.clear()
        password_input.fill(AMAZON_PASSWORD)
        time.sleep(0.5)
        print("[SUCCESS] Entered password")
        
        # Click sign in
        signin_btn = find_first_visible(page, SIGNIN_SELECTORS)
        if not signin_btn:
            raise RuntimeError("Could not find sign-in button")
        
        print("[INFO] Clicking sign-in button...")
        human_click(signin_btn, delay_after=0.5)
        wait_for_page_load(page)
        time.sleep(3)
        print("[SUCCESS] Sign-in button clicked")
        
        current_url = page.url
        print(f"[DEBUG] Current URL: {current_url}")
        
        # Check for security verification
        if "cvf/approval" in current_url or "cvf/verify" in current_url:
            print("\n" + "="*60)
            print("[WARNING] AMAZON SECURITY VERIFICATION DETECTED")
            print("="*60)
            print("Waiting for verification...")
            for wait_attempt in range(24):
                time.sleep(5)
                current_url = page.url
                if "cvf/approval" not in current_url and "cvf/verify" not in current_url:
                    print(f"[SUCCESS] Verification completed")
                    break
                if wait_attempt % 6 == 0:
                    print(f"[INFO] Still waiting... ({wait_attempt * 5}s)")
            else:
                print("[WARNING] Verification timeout - continuing...")
            time.sleep(2)
            current_url = page.url
        
        # Check for OTP
        print("\n[4/5] Checking for two-factor authentication...")
        otp_input = None
        for retry in range(3):
            otp_input = find_first_visible(page, OTP_SELECTORS, timeout=3000)
            if otp_input:
                print(f"[SUCCESS] OTP input field found")
                break
            else:
                if retry < 2:
                    print(f"[INFO] Waiting for OTP field... (attempt {retry + 1}/3)")
                    time.sleep(3)
                else:
                    print("[INFO] OTP not required")
        
        if otp_input:
            print("[INFO] Two-factor authentication required")
            print("[INFO] Waiting 5 seconds for OTP email...")
            time.sleep(5)
            
            print("[INFO] Retrieving OTP from Gmail...")
            otp_code = get_amazon_otp_from_gmail(max_age_minutes=10, max_retries=20, retry_delay=5)
            
            if not otp_code:
                print("\n[WARNING] OTP not found, retrying...")
                time.sleep(30)
                otp_code = get_amazon_otp_from_gmail(max_age_minutes=10, max_retries=10, retry_delay=5)
                
                if not otp_code:
                    print("\n" + "="*60)
                    print("[ERROR] AUTOMATIC OTP RETRIEVAL FAILED")
                    print("="*60)
                    raise RuntimeError("Failed to retrieve OTP code")
            
            print(f"\n[INFO] Entering OTP: {otp_code}")
            otp_input.clear()
            otp_input.fill(otp_code)
            time.sleep(0.5)
            print(f"[SUCCESS] OTP entered")
            
            print("[INFO] Submitting OTP...")
            otp_submit = find_first_visible(page, OTP_SUBMIT_SELECTORS, timeout=5000)
            if otp_submit:
                human_click(otp_submit, delay_after=1.0)
                wait_for_page_load(page)
                time.sleep(3)
                print("[SUCCESS] OTP submitted")
            else:
                raise RuntimeError("Could not find OTP submit button")
        else:
            print("[INFO] No two-factor authentication required")
        
        # Verify login
        print("\n[5/5] Verifying login...")
        current_url = page.url
        
        if "ap/signin" in current_url or "ap/cvf" in current_url:
            print(f"\n[ERROR] Login failed")
            print(f"[DEBUG] Current URL: {current_url}")
            return False
        
        print(f"[SUCCESS] Login successful!")
        print(f"[DEBUG] Current URL: {current_url}")
        
        # Save session
        print(f"\n[INFO] Saving session to {SESSION_FILE}...")
        context.storage_state(path=SESSION_FILE)
        print(f"[SUCCESS] Session saved")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Login failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def check_session_valid(page):
    """
    Check if saved session is still valid
    
    Args:
        page: Playwright page object
        
    Returns:
        True if session is valid, False otherwise
    """
    try:
        page.goto("https://www.amazon.co.jp/", wait_until="domcontentloaded", timeout=10000)
        time.sleep(2)
        
        current_url = page.url
        if "ap/signin" not in current_url and "ap/cvf" not in current_url:
            try:
                account_nav = page.locator("#nav-link-accountList")
                if account_nav.count() > 0:
                    return True
            except:
                pass
            
            if "amazon.co.jp" in current_url and "/ap/" not in current_url:
                return True
        
        return False
    except Exception as e:
        print(f"[WARNING] Could not verify session: {e}")
        return False


def run_category_search():
    """
    Main automation workflow for category search:
    1. Login to Amazon
    2. Navigate to Business Discounts page
    3. Search for each category keyword
    4. Display all products with pagination for each keyword
    """
    print("\n" + "="*70)
    print(" "*10 + "AMAZON CATEGORY SEARCH AUTOMATION")
    print("="*70 + "\n")

    with sync_playwright() as p:
        try:
            print("[INIT] Launching Chrome browser (Japanese locale)...")
            browser = p.chromium.launch(
                channel="chrome",
                headless=False,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--start-maximized",
                    "--lang=ja-JP",
                ],
            )

            session_path = Path(SESSION_FILE)
            storage_state = None
            if session_path.exists():
                try:
                    raw = session_path.read_text(encoding="utf-8", errors="ignore").strip()
                    if not raw:
                        raise ValueError("session file is empty")
                    json.loads(raw)
                    storage_state = str(session_path)
                    print(f"[INFO] Loaded saved session: {SESSION_FILE}")
                except Exception as e:
                    print(f"[WARNING] Invalid session file: {e}")
                    try:
                        session_path.unlink()
                        print(f"[INFO] Deleted invalid session file")
                    except Exception:
                        pass

            context = browser.new_context(
                no_viewport=True,
                locale="ja-JP",
                timezone_id="Asia/Tokyo",
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                storage_state=storage_state,
            )
            page = context.new_page()
            print("[SUCCESS] Browser launched\n")

            # Login or use saved session
            if storage_state and check_session_valid(page):
                print("[SUCCESS] Using saved Amazon session")
                login_success = True
            else:
                if session_path.exists():
                    print("[INFO] Session expired, logging in again")
                else:
                    print("[INFO] No saved session, logging in")
                login_success = login_to_amazon(page, context)

            if not login_success:
                print("\n[ERROR] Login failed. Closing browser...")
                time.sleep(5)
                browser.close()
                return False

            # Navigate to Business Discounts page
            print("\n" + "="*60)
            print("STEP 2: NAVIGATING TO BUSINESS DISCOUNTS")
            print("="*60)
            page.goto(BUSINESS_DISCOUNTS_URL, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
            time.sleep(3)
            print("[SUCCESS] Business Discounts page loaded")

            # Search for each keyword
            print("\n" + "="*60)
            print("STEP 3: SEARCHING CATEGORIES")
            print("="*60)
            print(f"[INFO] Will search for {len(SEARCH_KEYWORDS)} categories:")
            for i, keyword in enumerate(SEARCH_KEYWORDS, 1):
                print(f"  {i}. {keyword}")
            print("="*60)

            for keyword_index, keyword in enumerate(SEARCH_KEYWORDS, 1):
                print(f"\n{'='*70}")
                print(f"CATEGORY {keyword_index}/{len(SEARCH_KEYWORDS)}: {keyword}")
                print(f"{'='*70}")
                
                success = search_and_display_products(page, keyword)
                
                if not success:
                    print(f"[WARNING] 検索失敗: {keyword}")
                else:
                    print(f"\n[SUCCESS] カテゴリー '{keyword}' の検索完了")
                
                # Longer pause before next search for clarity
                if keyword_index < len(SEARCH_KEYWORDS):
                    print(f"\n[INFO] 次の検索まで5秒待機します...")
                    time.sleep(5)  # Increased from 3 to 5 seconds

            # All searches completed
            print("\n" + "="*70)
            print(" "*15 + "✓ 全ての検索が完了しました ✓")
            print("="*70)
            print(f"[SUCCESS] {len(SEARCH_KEYWORDS)} カテゴリーの検索完了")
            print("="*70)

            print("\n[INFO] ブラウザを15秒間開いたままにします（結果確認用）...")
            time.sleep(15)  # Increased from 10 to 15 seconds
            
            print("\n[INFO] Closing browser...")
            browser.close()
            print("[SUCCESS] Browser closed")
            
            return True

        except Exception as e:
            print(f"\n[ERROR] Automation failed: {e}")
            import traceback
            traceback.print_exc()
            return False


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    
    print("\n" + "="*70)
    print(" "*10 + "AMAZON CATEGORY SEARCH TOOL")
    print("="*70)
    print(f"\nConfiguration:")
    print(f"  - Email: {AMAZON_EMAIL}")
    print(f"  - Search Keywords:")
    for i, keyword in enumerate(SEARCH_KEYWORDS, 1):
        print(f"    {i}. {keyword}")
    print(f"  - Session File: {SESSION_FILE}")
    print("="*70)
    print("\nExecution Flow:")
    print("  1. Login to Amazon (or use saved session)")
    print("  2. Navigate to Business Discounts page")
    print("  3. Search for each category keyword")
    print("  4. Display all products (with pagination)")
    print("="*70 + "\n")
    
    # Check if Gmail credentials exist
    if not GMAIL_CREDENTIALS_FILE.exists():
        print("[ERROR] Gmail API credentials file not found!")
        print(f"Expected location: {GMAIL_CREDENTIALS_FILE}")
        print("\nPlease download OAuth 2.0 credentials and save to 'data' folder")
        return False
    
    # Run automation
    success = run_category_search()
    
    if success:
        print("\n" + "="*70)
        print(" "*20 + "COMPLETED SUCCESSFULLY!")
        print("="*70 + "\n")
        return True
    else:
        print("\n" + "="*70)
        print(" "*20 + "AUTOMATION FAILED")
        print("="*70 + "\n")
        return False


if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)
