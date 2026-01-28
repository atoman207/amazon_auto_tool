#!/usr/bin/env python3
"""
Amazon Business Automation - Complete Solution
Automatically logs in, retrieves OTP from Gmail, applies filters, and displays products
"""

import os
import re
import base64
import time
import sys
import json
from pathlib import Path
from datetime import datetime, timedelta
from html.parser import HTMLParser
from html import unescape

# Check and import required packages with helpful error messages
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

try:
    import gspread
except ImportError as e:
    print("\n" + "="*70)
    print("[ERROR] gspread is not installed!")
    print("="*70)
    print("Please install it by running:")
    print("  pip install gspread")
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
GMAIL_CREDENTIALS_FILE = Path('data/client_secret_446842116198-nke8rjis6iaeuagepsp9p5gvbsu2cte4.apps.googleusercontent.com.json')
GMAIL_TOKEN_FILE = Path('token.json')

# Google Sheets API settings
SHEETS_SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]
SHEETS_CREDENTIALS_FILE = GMAIL_CREDENTIALS_FILE  # Use same OAuth credentials
SHEETS_TOKEN_FILE = Path('sheets_token.json')
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1t_HjbOjlcgwZACo2glY8w-OfUVAGa3TEcX2h5wIwejk/edit?hl=ja&gid=0#gid=0"
SPREADSHEET_ID = "1t_HjbOjlcgwZACo2glY8w-OfUVAGa3TEcX2h5wIwejk"  # Extracted from URL

# Session file
SESSION_FILE = "amazon_session.json"

# Amazon URLs (Japanese site)
AMAZON_LOGIN_URL = "https://www.amazon.co.jp/ap/signin?openid.pape.max_auth_age=900&openid.return_to=https%3A%2F%2Fwww.amazon.co.jp%2Fgp%2Fyourstore%2Fhome%3Fpath%3D%252Fgp%252Fyourstore%252Fhome%26signIn%3D1%26useRedirectOnSuccess%3D1%26action%3Dsign-out%26ref_%3Dabn_yadd_sign_out&openid.assoc_handle=jpflex&openid.mode=checkid_setup&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
BUSINESS_DISCOUNTS_URL = "https://www.amazon.co.jp/ab/business-discounts?ref_=abn_cs_savings_guide&pd_rd_r=242d4956-5f68-4e46-bc0d-0fc896eaadf4&pd_rd_w=jg2kX&pd_rd_wg=rMzwy"

# Filter settings
SELECTED_CATEGORIES = [
    "IT関連機器",           # IT Equipment
    "医療用品・消耗品",     # Medical Supplies
    "日用品・食品・飲料"    # Daily Necessities
]
MIN_DISCOUNT_PERCENT = 5

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

# Selectors for filtering (exact XPaths provided by user)
CATEGORY_DROPDOWN_BUTTON = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[1]/div[1]/span/span/input"

# Category checkboxes (IT関連機器, 医療用品・消耗品, 日用品・食品・飲料)
# Clicking the div container which contains the label and checkbox
CATEGORY_IT_EQUIPMENT = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[1]/div[2]/div[2]/fieldset/div[3]"
CATEGORY_MEDICAL_SUPPLIES = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[1]/div[2]/div[2]/fieldset/div[5]"
CATEGORY_DAILY_NECESSITIES = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[1]/div[2]/div[2]/fieldset/div[7]"

# Show results button after category selection
CATEGORY_SHOW_RESULTS_BUTTON = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[1]/div[2]/div[3]/div[2]/span/span"

# Discount dropdown and filter
DISCOUNT_DROPDOWN_BUTTON = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[2]/div[1]/span/span/input"

# Clicking the div container for 5% discount
DISCOUNT_5_PERCENT_RADIO = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[2]/div[2]/div[2]/fieldset/div[1]"
DISCOUNT_SHOW_RESULTS_BUTTON = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/div[2]/div[2]/div[4]/div[2]/span/span/input"

# Sort dropdown and option (ビジネス割引: 降順 = Business Discount: Descending)
SORT_DROPDOWN_BUTTON = "xpath=/html/body/div[1]/div[1]/div/div/div[3]/section/div/div/div/div/div[1]/div[2]/span/span/span/span/span/span[1]"
SORT_BUSINESS_DISCOUNT_DESC = "xpath=/html/body/div[3]/div/div/ul/li[3]/a"


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
    Extract 6-digit OTP code from HTML email using XPath-like structure
    The OTP is in: /html/body/div[6]/.../table/tbody/tr[4]/td/div/span
    
    Args:
        html_text: Email HTML body
    
    Returns:
        6-digit OTP code as string, or None if not found
    """
    if not html_text:
        return None
    
    try:
        # Method 1: Try to find the specific table structure from XPath
        # Look for table > tbody > tr[4] > td > div > span pattern
        # This matches the XPath: /table/tbody/tr[4]/td/div/span
        table_pattern = r'<table[^>]*>.*?<tbody[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<tr[^>]*>.*?<td[^>]*>.*?<div[^>]*>.*?<span[^>]*>(\d{6})</span>'
        match = re.search(table_pattern, html_text, re.DOTALL | re.IGNORECASE)
        if match:
            otp = match.group(1)
            if len(otp) == 6 and otp.isdigit():
                return otp
        
        # Method 2: Find all spans with 6-digit numbers and check context
        span_pattern = r'<span[^>]*>(\d{6})</span>'
        spans = re.findall(span_pattern, html_text, re.IGNORECASE)
        for span_text in spans:
            if len(span_text) == 6 and span_text.isdigit():
                # Check if it's in a table structure (more likely to be OTP)
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
    Extract 6-digit OTP code from email text (plain text or HTML)
    
    Args:
        text: Email body text (plain or HTML)
    
    Returns:
        6-digit OTP code as string, or None if not found
    """
    if not text:
        return None
    
    # First try HTML extraction (for the specific XPath structure)
    if '<html' in text.lower() or '<body' in text.lower() or '<table' in text.lower():
        html_otp = extract_otp_from_html(text)
        if html_otp:
            return html_otp
    
    # Then try regex patterns for plain text
    patterns = [
        r'確認コード(?:は|:|：)(?:次のとおりです)?(?:\s*[:：]\s*)?(\d{6})',
        r'verification\s+code(?:\s+is)?(?:\s*[:：]\s*)?(\d{6})',
        r'コード(?:\s*[:：]\s*)(\d{6})',
        r'(?:^|\s)(\d{6})(?:\s|$)',  # Standalone 6-digit number
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
    Prioritizes HTML over plain text for better OTP extraction
    
    Args:
        payload: Message payload from Gmail API
    
    Returns:
        Decoded email body text (HTML preferred)
    """
    html_body = ""
    text_body = ""
    
    # Check if body is in payload.body.data
    if 'body' in payload and 'data' in payload['body']:
        body_data = payload['body']['data']
        decoded = base64.urlsafe_b64decode(body_data).decode('utf-8', errors='ignore')
        # Check if it's HTML
        if '<html' in decoded.lower() or '<body' in decoded.lower() or '<table' in decoded.lower():
            html_body = decoded
        else:
            text_body = decoded
    
    # Check if body is in parts (multipart email)
    if 'parts' in payload:
        for part in payload['parts']:
            mime_type = part.get('mimeType', '')
            
            # Recursively check nested parts
            if 'parts' in part:
                nested_result = decode_email_body(part)
                if '<html' in nested_result.lower() or '<table' in nested_result.lower():
                    html_body += nested_result
                else:
                    text_body += nested_result
            
            # Get text/html (preferred for OTP extraction)
            elif mime_type == 'text/html':
                if 'data' in part['body']:
                    part_data = part['body']['data']
                    decoded = base64.urlsafe_b64decode(part_data).decode('utf-8', errors='ignore')
                    html_body += decoded + "\n"
            
            # Get text/plain (fallback)
            elif mime_type == 'text/plain':
                if 'data' in part['body']:
                    part_data = part['body']['data']
                    decoded = base64.urlsafe_b64decode(part_data).decode('utf-8', errors='ignore')
                    text_body += decoded + "\n"
    
    # Return HTML if available, otherwise plain text
    return html_body if html_body else text_body


def get_amazon_otp_from_gmail(max_age_minutes=5, max_retries=12, retry_delay=5):
    """
    Get latest Amazon OTP code from Gmail
    
    Args:
        max_age_minutes: Only check emails from last N minutes (default: 5)
        max_retries: Maximum number of retry attempts (default: 12)
        retry_delay: Seconds to wait between retries (default: 5)
    
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
    
    # Calculate timestamp for search query (Unix timestamp)
    since_time = int((datetime.now() - timedelta(minutes=max_age_minutes)).timestamp())
    
    # Search query for Amazon verification emails - keep it broad.
    # Amazon can send from various addresses and subjects can vary.
    query = (
        f'(from:amazon.co.jp OR from:account-update@amazon.co.jp OR from:no-reply@amazon.co.jp '
        f'OR from:auto-confirm@amazon.co.jp) after:{since_time}'
    )
    
    print(f"\n[INFO] Searching for Amazon verification emails from last {max_age_minutes} minutes...")
    print(f"[INFO] Query: {query}\n")
    
    for attempt in range(1, max_retries + 1):
        try:
            # Search for messages (newest first)
            results = service.users().messages().list(
                userId='me',
                q=query,
                maxResults=10  # Get last 10 messages
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
            
            # Check each message for OTP (newest first)
            for idx, message in enumerate(messages, 1):
                try:
                    # Get full message
                    msg = service.users().messages().get(
                        userId='me',
                        id=message['id'],
                        format='full'
                    ).execute()
                    
                    # Get subject and from
                    headers = msg['payload'].get('headers', [])
                    subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'No Subject')
                    from_email = next((h['value'] for h in headers if h['name'].lower() == 'from'), 'Unknown')
                    
                    print(f"\n  Checking Email {idx}:")
                    print(f"    From: {from_email[:50]}")
                    print(f"    Subject: {subject[:60]}...")
                    
                    # Optional: reject old emails by internalDate if present
                    try:
                        internal_ms = int(msg.get("internalDate", "0"))
                        if internal_ms:
                            age_seconds = (time.time() - (internal_ms / 1000.0))
                            if age_seconds > (max_age_minutes * 60):
                                print("    [INFO] Skipping (too old)")
                                continue
                    except Exception:
                        pass

                    # Decode email body (HTML preferred)
                    body_text = decode_email_body(msg['payload'])
                    
                    # Extract OTP
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
            
            # If we checked all messages and no OTP found, wait and retry
            if attempt < max_retries:
                print(f"\n[Attempt {attempt}/{max_retries}] OTP not found in existing emails, waiting {retry_delay}s for new email...")
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
# GOOGLE SHEETS API FUNCTIONS
# ============================================================================

def get_sheets_service():
    """
    Authenticate and return Google Sheets API service using gspread
    
    Returns:
        gspread client object
    """
    creds = None
    
    # Load existing token if available
    if SHEETS_TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(SHEETS_TOKEN_FILE), SHEETS_SCOPES)
        except Exception as e:
            print(f"[WARNING] Could not load Sheets token: {e}")
            creds = None
    
    # If no valid credentials, do OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("[INFO] Refreshing expired Sheets token...")
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"[WARNING] Token refresh failed: {e}")
                creds = None
        
        if not creds:
            if not SHEETS_CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"\n[ERROR] Google credentials file not found: {SHEETS_CREDENTIALS_FILE}\n"
                    "Please download OAuth credentials from Google Cloud Console."
                )
            
            print("\n" + "="*60)
            print("GOOGLE SHEETS API AUTHORIZATION REQUIRED")
            print("="*60)
            print("A browser window will open for Google authorization.")
            print("="*60 + "\n")
            
            flow = InstalledAppFlow.from_client_secrets_file(
                str(SHEETS_CREDENTIALS_FILE), SHEETS_SCOPES
            )
            creds = flow.run_local_server(port=0)
            
            # Save credentials for future use
            SHEETS_TOKEN_FILE.write_text(creds.to_json())
            print(f"[SUCCESS] Sheets token saved to {SHEETS_TOKEN_FILE}")
    
    # Return gspread client
    return gspread.authorize(creds)


def initialize_google_sheets():
    """
    Initialize Google Sheets connection and ensure headers are present
    
    Returns:
        Tuple of (gspread_client, worksheet, current_row_number) or (None, None, None) if failed
    """
    try:
        # Authenticate with Google Sheets
        gc = get_sheets_service()
        
        # Open the spreadsheet
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.sheet1  # Use first sheet
        
        # Read existing data to find the last row number
        existing_data = worksheet.get_all_values()
        
        # Prepare header row (Japanese to match existing spreadsheet)
        headers = [
            "No",
            "created_time",
            "ASIN",
            "商品名",
            "商品数",
            "参考価格",
            "数量別価格 （円）",
            "割引率（％）",
            "割引額（円）"
        ]
        
        # Determine starting row number
        if len(existing_data) == 0:
            # No data at all - write headers and start from 1
            worksheet.update('A1', [headers])
            current_number = 1
        elif len(existing_data) == 1:
            # Only headers exist - start from 1
            # Update headers if they don't match
            if existing_data[0] != headers:
                worksheet.update('A1', [headers])
            current_number = 1
        else:
            # Data exists - find the last number and continue from there
            # Update headers if they don't match
            if existing_data[0] != headers:
                worksheet.update('A1', [headers])
            
            # Find the last row number
            last_row_data = existing_data[-1]
            try:
                # Try to get the number from the first column
                last_number = int(last_row_data[0]) if last_row_data[0] else 0
            except (ValueError, IndexError):
                # If we can't parse it, count the rows
                last_number = len(existing_data) - 1
            
            current_number = last_number + 1
        
        return gc, worksheet, current_number
        
    except Exception as e:
        print(f"\n[ERROR] Failed to initialize Google Sheets: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None


def append_product_to_sheets(worksheet, product_rows, current_number):
    """
    Append a single product (with all its quantity tiers) to Google Sheets immediately
    
    Args:
        worksheet: gspread worksheet object
        product_rows: List of dictionaries for this product (one per quantity tier)
        current_number: Current sequential product number
        
    Returns:
        Next product number to use, or None if failed
    """
    try:
        rows = []
        
        for idx, product in enumerate(product_rows):
            # Check if this is the first tier of the product
            is_first_tier = product.get('is_first_tier', idx == 0)
            
            row = [
                current_number if is_first_tier else '',  # Sequential number (only for first tier)
                product.get('created_time', ''),  # Timestamp (already blank for non-first tiers)
                product.get('asin', ''),  # ASIN (already blank for non-first tiers)
                product.get('name', ''),  # Name (already blank for non-first tiers)
                product.get('quantity', ''),
                product.get('reference_price', ''),
                product.get('unit_price', ''),
                product.get('discount_rate', ''),
                product.get('discount_amount', '')
            ]
            rows.append(row)
        
        # Append rows for this product
        worksheet.append_rows(rows)
        
        # Return next number (increment only once per product, not per tier)
        return current_number + 1
        
    except Exception as e:
        print(f"    [ERROR] Failed to append to Google Sheets: {e}")
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
    Click with basic human-like mouse movement.
    Locator must resolve to exactly one element (use .first where needed).
    """
    try:
        locator.scroll_into_view_if_needed(timeout=TIMEOUT_MS)
    except Exception:
        pass
    try:
        box = locator.bounding_box()
        if box:
            x = box["x"] + min(10, box["width"] / 2)
            y = box["y"] + min(10, box["height"] / 2)
            locator.page.mouse.move(x, y)
            time.sleep(0.1)
    except Exception:
        pass
    locator.click()
    time.sleep(delay_after)


def wait_for_page_load(page, timeout=TIMEOUT_MS):
    """Wait for page to load"""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout)
    except PWTimeoutError:
        print("[WARNING] Page load timeout, continuing...")


def scroll_products_page(page, scroll_times=10, scroll_delay=1.0):
    """
    Slowly scroll down the products page to display all filtered products
    
    Args:
        page: Playwright page object
        scroll_times: Number of times to scroll (default: 10)
        scroll_delay: Delay between scrolls in seconds (default: 1.0)
    """
    print("\n[INFO] Scrolling through products page...")
    try:
        for i in range(scroll_times):
            # Scroll down 500 pixels
            page.mouse.wheel(0, 500)
            time.sleep(scroll_delay)
            print(f"[INFO] Scrolled {i + 1}/{scroll_times} times")
        print("[SUCCESS] Finished scrolling through products")
    except Exception as e:
        print(f"[WARNING] Error during scrolling: {e}")


def extract_number(text):
    """Extract numeric value from text (handles Japanese currency format)"""
    if not text:
        return None
    # Remove currency symbols, commas, and extract number
    cleaned = re.sub(r'[¥,円JPY\s]', '', text)
    match = re.search(r'\d+(?:\.\d+)?', cleaned)
    return match.group(0) if match else None


def highlight_product_in_browser(page, container, asin, product_name=""):
    """
    Highlight the current product being scraped in the browser for visual feedback
    Shows green highlight and "SCRAPING..." label on the product
    
    Args:
        page: Playwright page object
        container: Product container element
        asin: Product ASIN for identification
        product_name: Product name for console logging (optional)
    """
    try:
        # Inject JavaScript to highlight this product and log to console
        page.evaluate(f"""
            (function() {{
                // Console logging for tracking
                console.log('%c🔄 SCRAPING PRODUCT', 'background: #00FF00; color: #000; font-size: 16px; font-weight: bold; padding: 5px;');
                console.log('ASIN: {asin}');
                console.log('Name: {product_name[:50] if product_name else "Loading..."}');
                console.log('─'.repeat(60));
                
                // Remove previous highlights
                document.querySelectorAll('.scraping-highlight').forEach(el => {{
                    el.classList.remove('scraping-highlight');
                    el.style.border = '';
                    el.style.backgroundColor = '';
                }});
                
                // Find and highlight current product
                const container = document.querySelector('[data-asin="{asin}"]')?.closest('.a-cardui, [data-a-card-type]');
                if (container) {{
                    container.classList.add('scraping-highlight');
                    container.style.border = '4px solid #00FF00';
                    container.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                    container.style.transition = 'all 0.3s ease';
                    container.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                    
                    // Add label showing it's being scraped
                    const label = document.createElement('div');
                    label.style.cssText = 'position: absolute; top: 5px; left: 5px; background: #00FF00; color: black; padding: 8px 12px; font-weight: bold; z-index: 9999; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,255,0,0.5); animation: pulse 1s infinite;';
                    label.innerHTML = '<span style="font-size: 14px;">🔄 SCRAPING ASIN: {asin}</span>';
                    label.className = 'scraping-label';
                    
                    // Add pulse animation
                    if (!document.getElementById('scraping-animation-style')) {{
                        const style = document.createElement('style');
                        style.id = 'scraping-animation-style';
                        style.textContent = `
                            @keyframes pulse {{
                                0%, 100% {{ transform: scale(1); }}
                                50% {{ transform: scale(1.05); }}
                            }}
                        `;
                        document.head.appendChild(style);
                    }}
                    
                    // Remove old label if exists
                    const oldLabel = document.querySelector('.scraping-label');
                    if (oldLabel) oldLabel.remove();
                    
                    // Make container relative if not already
                    if (getComputedStyle(container).position === 'static') {{
                        container.style.position = 'relative';
                    }}
                    
                    container.appendChild(label);
                    
                    // Change to "COMPLETE" after scraping
                    setTimeout(() => {{
                        if (label.parentNode) {{
                            label.style.background = '#32CD32';
                            label.innerHTML = '<span style="font-size: 14px;">✅ COMPLETE</span>';
                        }}
                        container.style.border = '2px solid #32CD32';
                        container.style.backgroundColor = 'rgba(50, 205, 50, 0.05)';
                    }}, 1500);
                    
                    // Remove label after showing complete
                    setTimeout(() => {{
                        if (label.parentNode) label.remove();
                    }}, 3000);
                }}
            }})();
        """)
        time.sleep(0.3)  # Brief pause to show highlight
    except Exception as e:
        # Don't fail scraping if highlight fails
        pass


def scrape_product_from_listing(container):
    """
    Scrape product details directly from listing page container
    Creates multiple rows for quantity-based pricing tiers
    Uses robust selectors with fallbacks
    
    Args:
        container: Playwright locator for product card container
        
    Returns:
        List of dictionaries (one per quantity tier), or empty list if failed
    """
    try:
        products_data = []
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # ===== Extract ASIN (CRITICAL) =====
        asin = ''
        try:
            # Try multiple ways to get ASIN
            asin_selectors = [
                '[data-asin]',
                '[data-asin]:not([data-asin=""])',
                'div[data-asin]',
                'section[data-asin]'
            ]
            for selector in asin_selectors:
                asin_elem = container.locator(selector).first
                if asin_elem.count() > 0:
                    asin = asin_elem.get_attribute('data-asin')
                    if asin and len(asin) == 10:  # ASIN is always 10 characters
                        break
        except Exception as e:
            print(f"    [DEBUG] ASIN extraction error: {e}")
        
        if not asin:
            return []  # Can't proceed without ASIN
        
        # ===== Extract Product Name (IMPORTANT) =====
        name = ''
        try:
            # Try multiple selectors for product name
            name_selectors = [
                'span.a-truncate-full.a-offscreen',
                '.a-truncate-full',
                'a[title]',  # Fallback: link title attribute
                'h2 a span'
            ]
            for selector in name_selectors:
                name_elem = container.locator(selector).first
                if name_elem.count() > 0:
                    name = name_elem.inner_text().strip()
                    if name:
                        break
            
            # Fallback: try getting from title attribute
            if not name:
                title_elem = container.locator('a[title]').first
                if title_elem.count() > 0:
                    name = title_elem.get_attribute('title')
        except Exception as e:
            print(f"    [DEBUG] Name extraction error: {e}")
        
        # ===== Extract Reference Price (Individual/Retail Price - 個人向け価格) =====
        reference_price = ''
        try:
            # Comprehensive selectors for reference price (strikethrough price)
            ref_selectors = [
                '._dmFsd_retailPriceMobileInt_22uHn .a-offscreen',  # Mobile view
                '._dmFsd_retailPriceInt_HVi7A .a-offscreen',  # Desktop view
                'span.a-price.a-text-price[data-a-strike="true"] .a-offscreen',
                '.a-text-price .a-offscreen',
                'span[data-a-strike="true"] .a-offscreen'
            ]
            for selector in ref_selectors:
                ref_elem = container.locator(selector).first
                if ref_elem.count() > 0:
                    ref_text = ref_elem.inner_text().strip()
                    reference_price = extract_number(ref_text)
                    if reference_price:
                        break
        except Exception as e:
            print(f"    [DEBUG] Reference price extraction error: {e}")
        
        # ===== Extract Base Discount Rate (from badge or savings text) =====
        discount_rate_base = ''
        try:
            discount_selectors = [
                'span._dmFsd_savingsBadge_25xkz',  # Badge at top
                'span._dmFsd_businessSavingsMobileInt_2V1aF',  # Mobile savings
                'div._dmFsd_businessSavingsInt_2W0Iq',  # Desktop savings
                'span:has-text("OFF")',  # Any span with "OFF" text
                'span:has-text("%")'  # Any span with percentage
            ]
            for selector in discount_selectors:
                discount_elem = container.locator(selector).first
                if discount_elem.count() > 0:
                    discount_text = discount_elem.inner_text().strip()
                    discount_rate_base = extract_number(discount_text)
                    if discount_rate_base:
                        break
        except Exception as e:
            print(f"    [DEBUG] Discount rate extraction error: {e}")
        
        # ===== Extract Quantity Tiers (KEY FEATURE) =====
        quantity_tiers = []
        try:
            # Find quantity picker items (hidden dropdown with data attributes)
            tier_items = container.locator('ul._dmFsd_qpDropdown_2UuXs li._dmFsd_qpItem_3tHmj').all()
            
            for tier_item in tier_items:
                try:
                    # Extract quantity (minimum quantity for this tier)
                    quantity = tier_item.get_attribute('data-minimum-quantity')
                    
                    # Extract price for this tier (numeric value without formatting)
                    tier_price = tier_item.get_attribute('data-numeric-value')
                    
                    if quantity and tier_price:
                        # Clean up the price value
                        tier_price_clean = tier_price.replace(',', '').replace('.00', '')
                        quantity_tiers.append({
                            'quantity': quantity,
                            'unit_price': tier_price_clean
                        })
                except Exception:
                    continue
        except Exception as e:
            print(f"    [DEBUG] Quantity tiers extraction error: {e}")
        
        # ===== Fallback: If no quantity tiers found, get base price =====
        if not quantity_tiers:
            base_price = ''
            try:
                base_price_selectors = [
                    'span.a-price._dmFsd_businessPriceMobileInt_3u3XJ .a-offscreen',  # Mobile business price
                    'span.a-price._dmFsd_businessPriceInt_oPUj8 .a-offscreen',  # Desktop business price
                    'span.a-price .a-offscreen:not([data-a-strike="true"])',  # Any non-strikethrough price
                    'span.a-price-whole'  # Price whole number
                ]
                for selector in base_price_selectors:
                    price_elem = container.locator(selector).first
                    if price_elem.count() > 0:
                        price_text = price_elem.inner_text().strip()
                        base_price = extract_number(price_text)
                        if base_price:
                            break
            except Exception as e:
                print(f"    [DEBUG] Base price extraction error: {e}")
            
            # Create single tier with quantity 1
            if base_price:
                quantity_tiers.append({
                    'quantity': '1',
                    'unit_price': base_price
                })
        
        # ===== Build Product Data Rows (one per quantity tier) =====
        for idx, tier in enumerate(quantity_tiers):
            # Only fill product info (timestamp, ASIN, name) for the FIRST tier
            # Subsequent tiers have these fields blank
            product_data = {
                'created_time': timestamp if idx == 0 else '',
                'asin': asin if idx == 0 else '',
                'name': name if idx == 0 else '',
                'quantity': tier['quantity'],
                'reference_price': reference_price,
                'unit_price': tier['unit_price'],
                'discount_rate': '',
                'discount_amount': '',
                'is_first_tier': idx == 0  # Flag to track first row for numbering
            }
            
            # Calculate discount rate and amount for this tier
            if reference_price and tier['unit_price']:
                try:
                    ref = float(reference_price)
                    curr = float(tier['unit_price'])
                    discount_amount = ref - curr
                    discount_rate = (discount_amount / ref) * 100 if ref > 0 else 0
                    product_data['discount_rate'] = f"{discount_rate:.1f}"
                    product_data['discount_amount'] = f"{discount_amount:.0f}"
                except Exception:
                    product_data['discount_rate'] = discount_rate_base if discount_rate_base else ''
                    product_data['discount_amount'] = ''
            else:
                product_data['discount_rate'] = discount_rate_base if discount_rate_base else ''
                product_data['discount_amount'] = ''
            
            products_data.append(product_data)
        
        return products_data
        
    except Exception as e:
        print(f"    [ERROR] Failed to scrape product from listing: {e}")
        import traceback
        traceback.print_exc()
        return []


def scrape_all_products(page, worksheet, current_number):
    """
    Scrape all products directly from the filtered results page
    Sends each product to Google Sheets IMMEDIATELY after scraping (real-time)
    Scrolls gradually and scrapes products as they appear (NO separate page opens)
    Creates multiple rows for products with quantity-based pricing tiers
    
    Args:
        page: Playwright page object
        worksheet: gspread worksheet object for real-time updates
        current_number: Starting product number for sequential numbering
        
    Returns:
        Number of unique products scraped
    """
    print("\n" + "="*60)
    print("STEP 4: SCRAPING & SENDING TO SHEETS (REAL-TIME)")
    print("="*60)
    
    try:
        scraped_asins = set()  # Track already scraped ASINs
        total_rows_sent = 0  # Track total rows sent
        scroll_count = 0
        no_new_products_count = 0  # Counter for consecutive scrolls with no new products
        max_consecutive_no_products = 5  # Stop after 5 scrolls with no new products
        
        print("\n[INFO] Starting real-time scrape-and-send process...")
        print("[INFO] Each product sent to Google Sheets immediately after scraping")
        print("[INFO] Products scraped directly from listing (no page opens)")
        print("[INFO] Multiple rows created for quantity-based pricing")
        print("[INFO] Will continue until no more products are found")
        print("="*60 + "\n")
        
        while True:  # Scrape until no more products found
            # Find currently visible product containers
            # Robust selector: try multiple patterns
            product_containers = page.locator("div.a-cardui._dmFsd_cardItem_1LFgv[data-a-card-type='basic']").all()
            
            # Fallback selector if primary doesn't work
            if len(product_containers) == 0:
                product_containers = page.locator("div.a-cardui._dmFsd_cardItem_1LFgv").all()
            
            if len(product_containers) == 0:
                print(f"[Scroll {scroll_count + 1}] No product containers found yet, scrolling...")
                page.mouse.wheel(0, 500)
                time.sleep(1.5)
                scroll_count += 1
                
                # Safeguard: don't scroll infinitely if page structure changed
                if scroll_count > 50:
                    print("\n[WARNING] Scrolled 50 times without finding products. Stopping.")
                    print("[INFO] Page structure may have changed. Please check selectors.")
                    break
                continue
            
            # Scrape new products from visible containers
            new_products_found = 0
            
            for container in product_containers:
                try:
                    # Check if this container has an ASIN (to verify it's a product)
                    asin_elem = container.locator('[data-asin]').first
                    if asin_elem.count() == 0:
                        continue
                    
                    asin = asin_elem.get_attribute('data-asin')
                    if not asin or asin in scraped_asins or len(asin) != 10:
                        continue
                    
                    # Mark as scraped
                    scraped_asins.add(asin)
                    
                    # Scrape product data first (to get name for highlight)
                    product_rows = scrape_product_from_listing(container)
                    
                    if product_rows:
                        # Get product name from first row
                        first_row = product_rows[0]
                        product_name = first_row.get('name', 'Unknown')
                        
                        # Highlight this product in the browser (visual feedback)
                        highlight_product_in_browser(page, container, asin, product_name)
                        
                        # Send to Google Sheets IMMEDIATELY
                        new_number = append_product_to_sheets(worksheet, product_rows, current_number)
                        
                        if new_number:
                            # Success!
                            new_products_found += 1
                            total_rows_sent += len(product_rows)
                            current_number = new_number
                            
                            # Show progress in terminal
                            print(f"  ✓ [{current_number - 1}] {asin} - {product_name[:50]}... ({len(product_rows)} tiers) → SENT")
                        else:
                            print(f"  ✗ Failed to send ASIN {asin} to sheets")
                    else:
                        print(f"  ⚠ No data extracted for ASIN {asin}")
                    
                except Exception as e:
                    print(f"  ✗ Error processing container: {e}")
                    continue
            
            # Check if we found new products in this scroll
            if new_products_found > 0:
                no_new_products_count = 0  # Reset counter
                print(f"\n[Scroll {scroll_count + 1}] Processed {new_products_found} new products")
                print(f"[INFO] Total: {len(scraped_asins)} products | {total_rows_sent} rows sent to sheets\n")
            else:
                no_new_products_count += 1
                print(f"[Scroll {scroll_count + 1}] No new products found")
                
                # If no new products found in consecutive scrolls, we've reached the end
                if no_new_products_count >= max_consecutive_no_products:
                    print(f"\n[INFO] No new products found after {max_consecutive_no_products} consecutive scrolls - reached end of results")
                    break
            
            # Scroll down to load more products
            print(f"[INFO] Scrolling down to load more products...")
            page.mouse.wheel(0, 500)
            time.sleep(1.5)  # Wait for lazy loading
            scroll_count += 1
        
        print("\n" + "="*60)
        print(f"[SUCCESS] Scraping & sending completed!")
        print(f"[INFO] Unique products: {len(scraped_asins)}")
        print(f"[INFO] Total rows sent: {total_rows_sent}")
        print(f"[INFO] Total scrolls: {scroll_count}")
        print(f"[INFO] View at: {SPREADSHEET_URL}")
        print("="*60)
        
        return len(scraped_asins)
        
    except Exception as e:
        print(f"\n[ERROR] Failed to scrape products: {e}")
        import traceback
        traceback.print_exc()
        return len(scraped_asins) if scraped_asins else 0


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
        # Navigate to login page (Japanese site)
        print("\n[1/5] Navigating to Amazon Japan login page...")
        # Use the specific login URL provided in configuration
        login_url = AMAZON_LOGIN_URL
        page.goto(login_url, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
        time.sleep(2)
        
        # Language check removed to respect the specific login URL parameters
        
        print("[SUCCESS] Loaded Amazon Japan login page")
        
        # Check for Passkey modal and close it if present
        print("\n[INFO] Checking for Passkey modal...")
        try:
            # Look for "Close" button (閉じる) in modal
            close_btn = page.locator('button:has-text("閉じる"), [aria-label="閉じる"], button:has-text("Close")').first
            if close_btn.is_visible(timeout=3000):
                print("[INFO] Passkey modal detected. Closing...")
                human_click(close_btn)
                time.sleep(1)
                print("[SUCCESS] Closed Passkey modal")
                
                # After closing, look for "Login" button if needed (as per instructions)
                # Usually closing the modal reveals the login form, but sometimes needs a click
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
        
        # Set up dialog handler BEFORE clicking continue to handle passkey alert
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
            time.sleep(1)  # Reduced wait time
            
        # Check if dialog was handled
        if dialog_handled:
            print("[SUCCESS] Passkey alert was automatically dismissed")
        else:
            print("[INFO] No passkey alert appeared")
        
        # MANUAL DISMISSAL APPROACH: Wait for user to manually close any modals
        print("\n" + "="*70)
        print("⚠️  PASSKEY MODAL MAY APPEAR - PLEASE CLOSE IT MANUALLY")
        print("="*70)
        print("[ACTION REQUIRED] If a passkey alert appears on the screen:")
        print("  1. Click the '閉じる' (Close) button to dismiss it")
        print("  2. The script will automatically continue once the modal is closed")
        print("="*70)
        
        
        # Wait for password field to become truly accessible
        print("\n[INFO] Waiting for password field to become accessible...")
        print("[INFO] (Script will auto-continue when modal is dismissed)")
        
        password_accessible = False
        max_wait_time = 120  # 2 minutes maximum wait
        check_interval = 1   # Check every 1 second
        elapsed_time = 0
        
        while not password_accessible and elapsed_time < max_wait_time:
            try:
                # Try to find and interact with password field
                password_test = find_first_visible(page, PASSWORD_SELECTORS, timeout=1000)
                
                if password_test:
                    # Test if we can actually type in it (not just find it)
                    try:
                        # Try to focus and test typing
                        password_test.focus(timeout=1000)
                        password_test.press_sequentially("", timeout=1000)  # Test typing capability
                        
                        # If we got here, field is truly accessible
                        password_accessible = True
                        print(f"\n[SUCCESS] Password field is accessible after {elapsed_time} seconds!")
                        break
                        
                    except Exception:
                        # Field exists but can't interact (modal blocking)
                        pass
                
                # Wait and check again
                time.sleep(check_interval)
                elapsed_time += check_interval
                
                # Print status every 5 seconds
                if elapsed_time % 5 == 0 and elapsed_time > 0:
                    print(f"[INFO] Still waiting... ({elapsed_time}s elapsed - please close any modals)")
                    
            except Exception as e:
                time.sleep(check_interval)
                elapsed_time += check_interval
        
        if password_accessible:
            print("[SUCCESS] Password field confirmed accessible - continuing automatically")
        else:
            print(f"[WARNING] Timeout after {max_wait_time}s - continuing anyway")
        
        # Remove dialog listener after handling
        try:
            page.remove_listener("dialog", handle_dialog)
        except:
            pass
        
        # Enter password (with retry after modal close)
        print("\n[3/5] Entering password...")
        password_input = None
        for retry in range(5):  # Try up to 5 times
            password_input = find_first_visible(page, PASSWORD_SELECTORS, timeout=3000)
            if password_input:
                print(f"[SUCCESS] Password field found (attempt {retry + 1}/5)")
                break
            else:
                if retry < 4:
                    print(f"[INFO] Password field not found yet, waiting 1 second... (attempt {retry + 1}/5)")
                    time.sleep(1)
                else:
                    raise RuntimeError("Could not find password input field after closing modal")
        
        if not password_input:
            raise RuntimeError("Could not find password input field")
        
        # Clear any existing text and enter password
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
        
        # Quick wait for page to process
        print("[INFO] Waiting for Amazon to process login...")
        time.sleep(3)  # Reduced wait time for faster execution
        print("[SUCCESS] Sign-in button clicked")
        
        
        # Check current URL for issues
        current_url = page.url
        print(f"[DEBUG] Current URL: {current_url}")
        
        # Check for CAPTCHA or security challenges
        if "cvf/approval" in current_url or "cvf/verify" in current_url:
            print("\n" + "="*60)
            print("[WARNING] AMAZON SECURITY VERIFICATION DETECTED")
            print("="*60)
            print("Amazon is asking for additional verification.")
            print("Waiting automatically for verification to complete...")
            print("="*60)
            # Wait automatically for user to complete verification (up to 2 minutes)
            print("[INFO] Waiting up to 2 minutes for verification...")
            for wait_attempt in range(24):  # 24 * 5 seconds = 2 minutes
                time.sleep(5)
                current_url = page.url
                if "cvf/approval" not in current_url and "cvf/verify" not in current_url:
                    print(f"[SUCCESS] Verification completed (waited {wait_attempt * 5} seconds)")
                    break
                if wait_attempt % 6 == 0:  # Print every 30 seconds
                    print(f"[INFO] Still waiting... ({wait_attempt * 5} seconds elapsed)")
            else:
                print("[WARNING] Verification timeout - continuing anyway...")
            time.sleep(2)
            current_url = page.url
        
        # Check for OTP page (quick check for faster execution)
        print("\n[4/5] Checking for two-factor authentication...")
        print("[INFO] Waiting for OTP input field to appear (up to 10 seconds)...")
        
        otp_input = None
        for retry in range(3):  # Try 3 times with 3 second intervals = 9 seconds total
            otp_input = find_first_visible(page, OTP_SELECTORS, timeout=3000)
            if otp_input:
                print(f"[SUCCESS] OTP input field found (attempt {retry + 1}/3)")
                break
            else:
                if retry < 2:
                    print(f"[INFO] OTP field not found yet, waiting 3 seconds... (attempt {retry + 1}/3)")
                    time.sleep(3)
                else:
                    print("[INFO] OTP input field not found after 10 seconds - assuming not required")
        
        if otp_input:
            print("[INFO] Two-factor authentication required")
            
            # Wait for email to arrive before searching
            print("[INFO] Waiting 5 seconds for OTP email to arrive...")
            time.sleep(5)
            
            # AUTOMATIC OTP RETRIEVAL with longer retry time and more attempts
            print("[INFO] Starting automatic OTP retrieval from Gmail...")
            otp_code = get_amazon_otp_from_gmail(max_age_minutes=10, max_retries=20, retry_delay=5)
            
            if not otp_code:
                # Retry with longer wait time
                print("\n[WARNING] OTP not found in first attempt")
                print("[INFO] Waiting additional 30 seconds and retrying...")
                time.sleep(30)
                otp_code = get_amazon_otp_from_gmail(max_age_minutes=10, max_retries=10, retry_delay=5)
                
                if not otp_code:
                    print("\n" + "="*60)
                    print("[ERROR] AUTOMATIC OTP RETRIEVAL FAILED")
                    print("="*60)
                    print("Could not retrieve OTP code automatically.")
                    print("Please check your Gmail inbox for the verification code.")
                    print("="*60)
                    raise RuntimeError("Failed to retrieve OTP code from Gmail")
            
            print(f"\n[INFO] Entering OTP: {otp_code}")
            # Clear any existing text and enter OTP code
            otp_input.clear()
            otp_input.fill(otp_code)
            time.sleep(0.5)
            print(f"[SUCCESS] OTP code entered: {otp_code}")
            
            # Submit OTP
            print("[INFO] Looking for OTP submit button...")
            otp_submit = find_first_visible(page, OTP_SUBMIT_SELECTORS, timeout=5000)
            if otp_submit:
                print("[INFO] Found OTP submit button. Clicking...")
                human_click(otp_submit, delay_after=1.0)
                wait_for_page_load(page)
                time.sleep(3)
                print("[SUCCESS] OTP submitted")
            else:
                print("[WARNING] OTP submit button not found")
                raise RuntimeError("Could not find OTP submit button")
        else:
            print("[INFO] No two-factor authentication required")
        
        # Verify login success
        print("\n[5/5] Verifying login...")
        current_url = page.url
        
        if "ap/signin" in current_url or "ap/cvf" in current_url:
            print(f"\n[ERROR] Login failed")
            print(f"[DEBUG] Current URL: {current_url}")
            print("\nPossible reasons:")
            print("- Incorrect password")
            print("- Incorrect OTP code")
            print("- CAPTCHA challenge")
            print("- Amazon security block")
            return False
        
        print(f"[SUCCESS] Login successful!")
        print(f"[DEBUG] Current URL: {current_url}")
        
        # Save session state
        print(f"\n[INFO] Saving session to {SESSION_FILE}...")
        context.storage_state(path=SESSION_FILE)
        print(f"[SUCCESS] Session saved")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Login failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def apply_filters_and_sort(page):
    """
    Apply category filters, discount filter, and sorting using exact XPaths
    
    Args:
        page: Playwright page object
        
    Returns:
        True if filters applied successfully, False otherwise
    """
    
    print("\n" + "="*60)
    print("STEP 3: APPLYING FILTERS AND SORTING")
    print("="*60)
    
    try:
        # ========== STEP 1: SELECT CATEGORIES ==========
        print("\n[1/5] Opening category dropdown...")
        
        category_dropdown = page.locator(CATEGORY_DROPDOWN_BUTTON)
        if category_dropdown.count() > 0:
            human_click(category_dropdown, delay_after=1.5)
            print("[SUCCESS] Category dropdown opened")
        else:
            print("[ERROR] Category dropdown not found")
            return False
        
        # Scroll down in the modal to view all categories
        print("\n[INFO] Scrolling within category modal to view all options...")
        try:
            # Wait a moment for modal to fully render
            time.sleep(1)
            
            # Scroll down within the modal using mouse wheel
            # Multiple scroll attempts to ensure we can see all categories
            for i in range(5):
                page.mouse.wheel(0, 200)  # Scroll down 200 pixels
                time.sleep(0.3)
            
            print("[SUCCESS] Scrolled through category modal")
            
            # Scroll back up a bit to ensure all three target categories are visible
            for i in range(2):
                page.mouse.wheel(0, -100)  # Scroll up 100 pixels
                time.sleep(0.2)
            
            time.sleep(0.5)
            print("[INFO] Category modal ready for selection")
            
        except Exception as e:
            print(f"[WARNING] Could not scroll modal: {e}")
        
        # Select the three categories using exact XPaths
        print("\n[2/5] Selecting 3 categories...")
        
        # Category 1: IT関連機器 (IT-related equipment)
        print("  [1/3] Selecting IT-related equipment...")
        it_checkbox = page.locator(CATEGORY_IT_EQUIPMENT)
        if it_checkbox.count() > 0:
            try:
                # Scroll element into view before clicking
                it_checkbox.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
                # Click the div container
                it_checkbox.click(force=True, timeout=3000)
                time.sleep(0.5)
                print("    [SUCCESS] IT-related equipment selected")
            except Exception as e:
                print(f"    [ERROR] Failed to select IT equipment: {e}")
        else:
            print("    [WARNING] IT-related equipment checkbox not found")
        
        # Category 2: 医療用品・消耗品 (Medical supplies and consumables)
        print("  [2/3] Selecting Medical supplies and consumables...")
        medical_checkbox = page.locator(CATEGORY_MEDICAL_SUPPLIES)
        if medical_checkbox.count() > 0:
            try:
                # Scroll element into view before clicking
                medical_checkbox.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
                # Click the div container
                medical_checkbox.click(force=True, timeout=3000)
                time.sleep(0.5)
                print("    [SUCCESS] Medical supplies selected")
            except Exception as e:
                print(f"    [ERROR] Failed to select medical supplies: {e}")
        else:
            print("    [WARNING] Medical supplies checkbox not found")
        
        # Category 3: 日用品・食品・飲料 (Daily necessities, food, and beverages)
        print("  [3/3] Selecting Daily necessities, food, and beverages...")
        daily_checkbox = page.locator(CATEGORY_DAILY_NECESSITIES)
        if daily_checkbox.count() > 0:
            try:
                # Scroll element into view before clicking
                daily_checkbox.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
                # Click the div container
                daily_checkbox.click(force=True, timeout=3000)
                time.sleep(0.5)
                print("    [SUCCESS] Daily necessities selected")
            except Exception as e:
                print(f"    [ERROR] Failed to select daily necessities: {e}")
        else:
            print("    [WARNING] Daily necessities checkbox not found")
        
        # Click "Show Results" button for categories
        print("\n[INFO] Clicking 'Show Results' button for categories...")
        category_show_results = page.locator(CATEGORY_SHOW_RESULTS_BUTTON)
        if category_show_results.count() > 0:
            try:
                category_show_results.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
            except:
                pass
            human_click(category_show_results, delay_after=2.0)
            print("[SUCCESS] Categories applied - waiting for page to reload...")
            time.sleep(2)  # Wait for page to process category filter
        else:
            print("[ERROR] Category 'Show Results' button not found")
            return False
        
        # ========== STEP 2: SELECT DISCOUNT FILTER ==========
        print("\n[3/5] Opening discount dropdown...")
        discount_dropdown = page.locator(DISCOUNT_DROPDOWN_BUTTON)
        if discount_dropdown.count() > 0:
            human_click(discount_dropdown, delay_after=1.5)
            print("[SUCCESS] Discount dropdown opened")
        else:
            print("[ERROR] Discount dropdown not found")
            return False
        
        # No need to scroll for 5% option as it's at the top
        print("\n[INFO] Discount modal opened, 5% option should be visible...")
        try:
            # Wait a moment for modal to fully render
            time.sleep(1)
            print("[SUCCESS] Discount modal ready")
            
        except Exception as e:
            print(f"[WARNING] Could not prepare discount modal: {e}")
        
        # Select 5% discount radio button
        print("\n[4/5] Selecting 5% discount filter...")
        discount_5_radio = page.locator(DISCOUNT_5_PERCENT_RADIO)
        if discount_5_radio.count() > 0:
            try:
                # Scroll element into view before clicking
                discount_5_radio.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
                # Click the div container for 5% discount
                discount_5_radio.click(force=True, timeout=3000)
                time.sleep(0.5)
                print("[SUCCESS] 5% discount filter selected")
            except Exception as e:
                print(f"[ERROR] Failed to select 5% discount: {e}")
        else:
            print("[WARNING] 5% discount radio button not found")
        
        # Click "Show Results" button for discount
        print("\n[INFO] Clicking 'Show Results' button for discount...")
        discount_show_results = page.locator(DISCOUNT_SHOW_RESULTS_BUTTON)
        if discount_show_results.count() > 0:
            try:
                discount_show_results.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
            except:
                pass
            human_click(discount_show_results, delay_after=2.0)
            print("[SUCCESS] Discount filter applied - waiting for page to reload...")
            time.sleep(2)  # Wait for page to process discount filter
        else:
            print("[ERROR] Discount 'Show Results' button not found")
            return False
        
        # ========== STEP 3: APPLY SORT ORDER ==========
        print("\n[5/5] Applying sort order (Business Discount: Descending)...")
        sort_button = page.locator(SORT_DROPDOWN_BUTTON)
        if sort_button.count() > 0:
            human_click(sort_button, delay_after=1.5)
            print("[SUCCESS] Sort dropdown opened")
            
            # Wait for sort menu to fully appear
            time.sleep(0.5)
            
            # Click "Business Discount: Descending" option
            sort_option = page.locator(SORT_BUSINESS_DISCOUNT_DESC)
            if sort_option.count() > 0:
                human_click(sort_option, delay_after=1.5)
                print("[SUCCESS] Sort order applied (Business Discount: Descending)")
            else:
                print("[WARNING] Business Discount: Descending option not found")
        else:
            print("[WARNING] Sort dropdown not found, skipping...")
        
        print("\n" + "="*60)
        print("[SUCCESS] ALL FILTERS AND SORTING APPLIED")
        print("="*60)
        print(f"  ✓ Categories: {', '.join(SELECTED_CATEGORIES)}")
        print(f"  ✓ Discount: {MIN_DISCOUNT_PERCENT}%+")
        print(f"  ✓ Sort: Business Discount (Descending)")
        print("="*60)
        print("\n[INFO] Products will be scraped while scrolling until all products are found")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Failed to apply filters: {e}")
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
        # If we're not on sign-in page, session is likely valid
        if "ap/signin" not in current_url and "ap/cvf" not in current_url:
            # Try to find account indicator
            try:
                account_nav = page.locator("#nav-link-accountList")
                if account_nav.count() > 0:
                    return True
            except:
                pass
            
            # If URL is amazon.co.jp homepage, likely logged in
            if "amazon.co.jp" in current_url and "/ap/" not in current_url:
                return True
        
        return False
    except Exception as e:
        print(f"[WARNING] Could not verify session: {e}")
        return False


def run_automation():
    """
    Main automation workflow (single visible browser):
    Milestone 1:
        1. Open Amazon (Japanese locale) using saved session if available
        2. If session invalid: enter email/password and trigger OTP
        3. Fetch OTP from Gmail API and enter automatically
        4. Save session
        5. Navigate to Business Discounts and apply clicks/filters
    Milestone 2:
        6. Scrape all product data (ASIN, name, prices, discount, etc.)
        7. Send scraped data to Google Sheets
    """
    print("\n" + "="*70)
    print(" "*10 + "AMAZON BUSINESS DISCOUNT AUTOMATION")
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
                    # Playwright expects valid JSON storage state. If file is empty/corrupt, ignore it.
                    raw = session_path.read_text(encoding="utf-8", errors="ignore").strip()
                    if not raw:
                        raise ValueError("session file is empty")
                    json.loads(raw)  # validate JSON
                    storage_state = str(session_path)
                    print(f"[INFO] Loaded saved session: {SESSION_FILE}")
                except Exception as e:
                    print(f"[WARNING] Ignoring invalid session file '{SESSION_FILE}': {e}")
                    try:
                        # Remove bad file so next successful login can save a clean session.
                        session_path.unlink()
                        print(f"[INFO] Deleted invalid session file '{SESSION_FILE}'.")
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

            if storage_state and check_session_valid(page):
                print("[SUCCESS] Using saved Amazon session.")
                login_success = True
            else:
                if session_path.exists():
                    print("[INFO] Saved session is missing/expired; logging in again.")
                else:
                    print("[INFO] No saved session; logging in.")
                login_success = login_to_amazon(page, context)

            if not login_success:
                print("\n[ERROR] Login failed. Closing browser in 5 seconds...")
                time.sleep(5)
                browser.close()
                return False

            print("\n" + "="*60)
            print("STEP: NAVIGATING TO BUSINESS DISCOUNTS")
            print("="*60)
            page.goto(BUSINESS_DISCOUNTS_URL, wait_until="domcontentloaded", timeout=TIMEOUT_MS)
            time.sleep(3)
            print("[SUCCESS] Business Discounts page loaded")

            apply_filters_and_sort(page)

            # MILESTONE 2: Initialize Google Sheets and scrape products (real-time sending)
            print("\n" + "="*60)
            print("STEP: INITIALIZING GOOGLE SHEETS")
            print("="*60)
            
            gc, worksheet, current_number = initialize_google_sheets()
            
            if not worksheet:
                print("\n[ERROR] Failed to initialize Google Sheets.")
                print("[INFO] Closing browser...")
                browser.close()
                return False
            
            print(f"[SUCCESS] Google Sheets initialized")
            print(f"[INFO] Starting product number: {current_number}")
            print(f"[INFO] Spreadsheet: {SPREADSHEET_URL}")
            
            # Scrape products and send to sheets in real-time
            print("\n" + "="*60)
            print("STEP: SCRAPING PRODUCTS (MILESTONE 2)")
            print("="*60)
            
            unique_products = scrape_all_products(page, worksheet, current_number)
            
            if unique_products == 0:
                print("\n[WARNING] No products were scraped.")
                print("[INFO] Closing browser...")
                browser.close()
                return False
            
            # Success message
            print("\n" + "="*70)
            print(" "*15 + "✓ ALL MILESTONES COMPLETED ✓")
            print("="*70)
            print(f"[SUCCESS] Scraped and sent {unique_products} products to Google Sheets")
            print(f"[INFO] Data sent in real-time (row by row)")
            print(f"[INFO] View at: {SPREADSHEET_URL}")
            print("="*70)

            # Close browser immediately after completion
            print("\n[INFO] Closing browser...")
            browser.close()
            print("[SUCCESS] Browser closed")
            
            return True  # Success if we got here

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
    print(" "*10 + "AMAZON BUSINESS DISCOUNT AUTOMATION TOOL")
    print("="*70)
    print(f"\nConfiguration:")
    print(f"  - Email: {AMAZON_EMAIL}")
    print(f"  - Categories: {', '.join(SELECTED_CATEGORIES)}")
    print(f"  - Min Discount: {MIN_DISCOUNT_PERCENT}%")
    print(f"  - Session File: {SESSION_FILE}")
    print("="*70)
    print("\nExecution Flow:")
    print("  Milestone 1:")
    print("    1. Gmail authorization (browser opens, then closes)")
    print("    2. Amazon login (browser reopens)")
    print("    3. Enter email and password")
    print("    4. Enter verification code from email")
    print("    5. Select product categories")
    print("    6. Display product list")
    print("  Milestone 2:")
    print("    7. Initialize Google Sheets connection")
    print("    8. Scrape products and send to Sheets IMMEDIATELY (real-time)")
    print("    9. Each product sent row-by-row as it's scraped")
    print("    10. Continue until no more products found")
    print("="*70 + "\n")
    
    # Check if Gmail credentials exist
    if not GMAIL_CREDENTIALS_FILE.exists():
        print("[ERROR] Gmail API credentials file not found!")
        print(f"Expected location: {GMAIL_CREDENTIALS_FILE}")
        print("\nPlease download your OAuth 2.0 credentials from:")
        print("https://console.cloud.google.com/apis/credentials")
        print("And save it to the 'data' folder.")
        return False
    
    # Run automation
    success = run_automation()
    
    if success:
        print("\n" + "="*70)
        print(" "*20 + "ALL STEPS COMPLETED SUCCESSFULLY!")
        print("="*70 + "\n")
        return True
    else:
        print("\n" + "="*70)
        print(" "*20 + "AUTOMATION FAILED")
        print("="*70)
        print("\n[ERROR] Please check the error messages above.")
        print("="*70 + "\n")
        return False


if __name__ == "__main__":
    import sys
    success = main()
    sys.exit(0 if success else 1)
