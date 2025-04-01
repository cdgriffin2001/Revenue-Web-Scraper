from bs4 import BeautifulSoup
import re
import time
import random
import pandas as pd
from difflib import SequenceMatcher
from collections import Counter, defaultdict
import numpy as np
from urllib.parse import quote_plus
import math 
from playwright.sync_api import sync_playwright


# ======================
# Configuration
# ======================
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15"
]

PROXIES = []  # Format: [{"server": "http://proxy:port", "username": "user", "password": "pass"}]

# ======================
# Browser Manager
# ======================

class BrowserManager:
    """Handles browser lifecycle with enhanced anti-detection features"""
    def __init__(self):
        self.playwright = None
        self.browser = None

    def __enter__(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-web-security",
                "--disable-site-isolation-trials"
            ],
            proxy=random.choice(PROXIES) if PROXIES else None
        )
        return self.browser

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

# ======================
# Revenue Parser
# ======================
def parse_revenue(text): 
    """Robust revenue parser with range handling and unit conversions"""
    # Handle ranges first
    range_pattern = r'\$?[\s]*([\d,\.]+)\s*-\s*\$?[\s]*([\d,\.]+)\s*(million|billion|m|b|k)\b'
    range_match = re.search(range_pattern, text, re.IGNORECASE)
    if range_match:
        try:
            val1 = float(range_match.group(1).replace(',', ''))
            val2 = float(range_match.group(2).replace(',', ''))
            unit = range_match.group(3).lower()[0]  # Take first letter
            multipliers = {'k': 0.001, 'm': 1, 'b': 1000}
            return min(val1, val2) * multipliers.get(unit, 1)
        except (ValueError, AttributeError):
            pass

    # Enhanced patterns list
    patterns = [
        (r'(?:USD|€|£)\s*([\d,\.]+)\s*([mbk])\b', 1, 2),
        (r'(?:~|approx\.?|about|around)\s*\$?\s*([\d,\.]+)\s*(m|b|k)il', 1, 2),
        (r'\$[\s]*([\d,\.]+)\s*(million|billion|m|b|k)\b', 1, 2),
        (r'(?:revenue|sales)[^\$]{0,20}\$?\s*([\d,\.]+)\s*(m|b|k)il', 1, 2),
        (r'\$[\s]*([\d,\.]+)\s*(?:USD|US D|Dollars?)\s*(m|b|k)il', 1, 2)
    ]
    
    # Rest of original logic with enhanced multiplier handling
    for pattern, val_group, unit_group in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                value = float(match.group(val_group).replace(',', ''))
                unit = (match.group(unit_group) or '')[0].lower()
                multipliers = {'k': 0.001, 'm': 1, 'b': 1000}
                return value * multipliers.get(unit, 1)
            except (ValueError, AttributeError):
                continue
    return None

# ======================
# Search Logic
# ======================

def find_company_revenue(company_name, company_address, revenue_site, city):
    """Enhanced search function with better stability and anti-blocking measures"""
    try:
        with BrowserManager() as browser:
            context = browser.new_context(
                user_agent=random.choice(USER_AGENTS),
                locale='en-US',
                timezone_id='America/Chicago',
                viewport={'width': 1366, 'height': 768},
                extra_http_headers={
                    'Accept-Language': 'en-US,en;q=0.9',
                    'Referer': 'https://www.bing.com/'
                }
            )
            
            page = context.new_page()
            
            try:
                # Improved initialization sequence
                page.goto("https://www.bing.com/", wait_until="networkidle")
                
                # Handle cookie consent if present
                try:
                    page.click('button#bnp_btn_accept', timeout=5000)
                except Exception:
                    pass

                # Build search query with URL encoding
                search_query = f"{company_name} revenue {company_address} {city} {revenue_site}"
                encoded_query = quote_plus(search_query)
                
                # Direct navigation instead of typing (more reliable)
                page.goto(f"https://www.bing.com/search?q={encoded_query}", wait_until="domcontentloaded")
                
                # Wait for core content with fallback
                try:
                    page.wait_for_selector('ol#b_results', timeout=15000)
                except Exception:
                    page.wait_for_selector('div.b_content', timeout=15000)

                # Human-like interactions
                for _ in range(3):  # Reduced but sufficient
                    page.mouse.move(
                        random.randint(300, 1000),  # More realistic movement range
                        random.randint(300, 500),
                        steps=random.randint(5, 10)
                    )
                    page.wait_for_timeout(random.randint(300, 800))

                # Enhanced content extraction
                content = page.inner_html('ol#b_results') or page.inner_html('body')
                soup = BeautifulSoup(content, 'html.parser')

                # Updated priority selectors (Bing 2025 structure)
                priority_selectors = [
                    '[data-tag*="revenue"]',# Elements with revenue metadata
                    '.b_caption.hasdl.b_stsp2',    # Div with multiple classes "b_caption hasdl b_stsp2"
                    '.b_lineclamp2',
                ]
                
                # parsing logic 
                for selector in priority_selectors:
                    elements = soup.select(selector)
                    for elem in elements:
                        text = elem.get_text(strip=True)
                        if 'revenue' in text.lower():
                            result = parse_revenue(text)
                            if result:
                                return (result, text)
                
                return None
                
            finally:
                context.close()
                
    except Exception as e:
        print(f"Search error: {str(e)}")
        return None

# ======================
# Main Execution
# ======================

def clean_text(text):
    """Enhanced cleaning with corporate suffix removal"""
    # Remove common corporate suffixes/prefixes
    suffixes = r'\b(inc|llc|ltd|corp|co|group|holdings|plc|limited|company|corporation|incorporated)\b'
    text = re.sub(suffixes, '', text, flags=re.IGNORECASE)
    
    # Standard cleaning
    return re.sub(r'[^\w\s]', '', text.lower()).strip()

def get_match_confidence(result_text, company_name, company_address):
    """Returns confidence score (0-2) with enhanced matching"""
    clean_result = clean_text(result_text)
    clean_name = clean_text(company_name)
    clean_address = clean_text(company_address) if company_address else ""

    # Name matching improvements
    name_in_result = clean_name in clean_result
    partial_ratio = SequenceMatcher(None, clean_name, clean_result).ratio()
    
    # Adjusted matching criteria
    name_match = any([
        name_in_result,
        partial_ratio > 0.85,
        len(clean_name.split()) == 1 and partial_ratio > 0.75
    ])

    # Address verification
    address_match = False
    if clean_address:
        address_parts = [p for p in clean_address.split() if len(p) > 3]
        result_parts = clean_result.split()
        matches = sum(1 for part in address_parts if part in result_parts)
        address_match = matches >= 2

    if name_match and address_match:
        return 2
    elif name_match or address_match:
        return 1
    return 0
def calculate_discrepancy(values):
    """Calculate coefficient of variation to measure spread"""
    if len(values) < 2:
        return 0
    std_dev = np.std(values)
    mean = np.mean(values)
    return (std_dev / mean) if mean != 0 else 0

def process_results(confidence_groups):
    """Enhanced decision logic with discrepancy checks"""
    for confidence in [2, 1, 0]:  # Priority order
        group = confidence_groups.get(confidence, [])
        if not group:
            continue

        # Convert to millions for comparison
        values = [v if v < 1e6 else v/1e6 for v in group]
        
        # Scenario: Single result handling
        if len(values) == 1:
            if confidence == 2:  # High confidence match
                return values[0]
            continue  # Skip lone medium/low confidence results

        # Calculate spread and consensus
        freq = Counter(values)
        max_freq = max(freq.values())
        common_values = [v for v, cnt in freq.items() if cnt == max_freq]
        discrepancy = calculate_discrepancy(values)

        # Scenario: Large discrepancy handling
        if discrepancy > 0.5:  
            return min(values)

        # Scenario: Consensus handling
        if max_freq >= 2:
            return min(common_values)  # Prefer conservative estimate

        # Scenario: No consensus but multiple values
        if len(values) >= 2:
            return int(np.median(values)) if discrepancy < 0.3 else min(values)

    return "N/A"

def revenue_web_scrape(company_name, company_loc, city):
    raw_results = []
    revenue_sites = ["zoom info", "rocket reach", "zippia", "duns & bradstreet", "estimate"]

    def process_scrape_call(name, address, site, city):
        result = find_company_revenue(name, address, site, city)
        return result or None

    # Primary search with location
    for site in revenue_sites:
        if result := process_scrape_call(company_name, company_loc, site, city):
            raw_results.append(result)

    # Fallback search if insufficient high-confidence results
    if len(raw_results) < 3:
        for site in revenue_sites:
            if result := process_scrape_call(company_name, "", site, ""):
                raw_results.append(result)

    # Result categorization
    confidence_groups = defaultdict(list)
    for value, text in raw_results:
        conf = get_match_confidence(text, company_name, company_loc)
        confidence_groups[conf].append(value)

    return process_results(confidence_groups)

# #####FOR TESTING############
# company_name = "CLOVER MEADOW LLC"
# company_loc = "23396 THOMPSON RD"
# city= "SHELL LAKE"

# test= revenue_web_scrape(company_name, company_loc,city)
# print(test)
# ###########################

# Load the data
file_path = 'Desktop/Test.xlsx'
data = pd.read_excel(file_path)
# print(data.head)
# print(data.columns)  # Check column names
# print(data['Adress'].head())  # Verify address data

def sanitize_value(value):
    if (
        value is None
        or value == "N/A"
        or (isinstance(value, float) and math.isnan(value))
    ):
        return ''
    return str(value)  # Convert to string if needed


import os
import pandas as pd

def get_processed_count():
    """Returns number of completed rows with atomic write safety"""
    if os.path.exists('Desktop/Updated_Test.xlsx'):
        try:
            existing = pd.read_excel('Desktop/Updated_Test.xlsx')
            return existing['Revenue(millions)'].notna().sum()
        except Exception as e:
            print(f"Error reading progress: {e}")
            return 0
    return 0

def valid_name(name):
    """Validate company names before processing"""
    if not name or pd.isna(name):
        return False
    return str(name).strip().lower() not in ['n/a', 'null', 'none']

#____________________Upload to Excel____________________________
# Iterate through each row in the DataFrame
try:
    # Load data with index reset for reliable iloc
    data = pd.read_excel('Desktop/Test.xlsx').reset_index(drop=True)
    processed = get_processed_count()
    
    print(f"Resuming from row {processed} of {len(data)}")

    # Modified processing loop with proper indentation
    for index, row in data.iloc[processed:].iterrows():
        revenue = None
        company_name = row.get('DBA NAME')
        company_loc = sanitize_value(row.get('ADDRESS'))
        city = sanitize_value(row.get('CITY'))
        backup_name = row.get('BUSINESS NAME')

        # Processing logic
        if valid_name(company_name):
            revenue = revenue_web_scrape(company_name, company_loc, city)
            print(f'Trying DBA Name: {company_name} - Revenue: {revenue}')

        if revenue is None and valid_name(backup_name):
            revenue = revenue_web_scrape(backup_name, company_loc, city)
            print(f'Fallback to BUSINESS NAME: {backup_name} - Revenue: {revenue}')

        data.at[index, 'Revenue(millions)'] = revenue if revenue else 'N/A'
        
        # Progress saving every 10 rows
        if (index + 1) % 10 == 0:
            temp_path = 'Desktop/Updated_Test_TEMP.xlsx'
            data.to_excel(temp_path, index=False)
            os.replace(temp_path, 'Desktop/Updated_Test.xlsx')
            print(f"Saved through row {index+1}/{len(data)}")

        # Rate limiting with progress awareness
        time.sleep(random.uniform(3.42, 7.54) * (1 + index/100))

    # Final save after completion
    data.to_excel('Desktop/Updated_Test.xlsx', index=False)
    print("Processing complete. Final results saved.")

except FileNotFoundError as e:
    print(f"Critical error: {str(e)}")
except Exception as e:
    print(f"Unexpected error: {str(e)}")
    # Attempt emergency save
    data.to_excel('Desktop/CRASH_RECOVERY.xlsx', index=False)
    print("Partial results saved to CRASH_RECOVERY.xlsx")

# ________________________________________________________________________________________________________________________________




