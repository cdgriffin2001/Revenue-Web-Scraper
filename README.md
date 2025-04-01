README

# Revenue Web Scraper

## Overview
The Revenue Web Scraper is a Python-based tool designed to automate the process of extracting company revenue data from online sources. It uses advanced web scraping techniques with Playwright for browser automation and BeautifulSoup for parsing HTML content. The tool supports batch processing of company data from Excel files and includes anti-detection measures to avoid being blocked by websites.

## Features
- **Browser Automation**: Uses Playwright to simulate human-like browsing behavior.
- **Revenue Parsing**: Extracts revenue information with support for ranges, unit conversions (e.g., millions, billions), and confidence scoring.
- **Batch Processing**: Reads company data from Excel files and updates results in real time.
- **Error Handling**: Includes robust error handling and crash recovery to save progress.
- **Anti-Detection Measures**: Implements techniques like randomized user agents, proxy support, and human-like mouse movements.

## Requirements
To run this project, you need:
- Python 3.8 or higher
- The following Python libraries:
  - `playwright`
  - `beautifulsoup4`
  - `pandas`
  - `numpy`
  - `openpyxl`

## Installation
1. Clone this repository to your local machine:

2. Install the required dependencies:
pip install -r requirements.txt


3. Set up Playwright (this is required for browser automation):
playwright install


## Usage
1. Prepare an input Excel file named `Test.xlsx` in the `Desktop/` directory. The file should contain columns like:
- `DBA NAME`: The company's doing-business-as name.
- `ADDRESS`: The company's address.
- `CITY`: The city where the company is located.
- `BUSINESS NAME`: A backup name for the company (if DBA name fails).

2. Run the scraper script:
python BSoup.py

3. The script will process each row in the Excel file and save the results to a new file named `Updated_Test.xlsx` in the same directory.

4. If the script crashes, it will save progress to a recovery file named `CRASH_RECOVERY.xlsx`.

## How It Works
1. **Search Logic**: The scraper uses Bing search to find relevant revenue information based on the company name, address, and city.
2. **Revenue Parsing**: Extracts revenue data using patterns like `$10M`, `$5 billion`, or ranges like `$1M-$5M`.
3. **Confidence Scoring**: Matches results against company names and addresses to ensure accuracy.
4. **Batch Processing**: Processes rows from an Excel file one by one, saving progress periodically.

## File Structure
