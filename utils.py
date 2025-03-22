from furl import furl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import re

def normalize_url(url):
    # Create a furl object and parse the URL
    f = furl(url)
    
    # Normalize domain name: convert to lowercase
    f.host = f.host.lower()

    # Remove fragment, normalize query parameters (optional), and keep case-sensitive path/query
    f.args = sorted(f.args.items())  # Optional: sort query parameters to avoid order issues
    f.path = str(f.path).rstrip('/')  # Remove trailing slashes for path comparison
    
    return f.url

def compare_urls(url1, url2):
    # Normalize both URLs and check equality
    return normalize_url(url1) == normalize_url(url2)
    
def InitDriver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument(f"--remote-debugging-port=9223")  # Attach to the existing browser window
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36"
    )
    options.add_argument("--headless")  # Run in headless mode (no window)
    options.add_argument("--disable-gpu")  # Disable GPU acceleration (useful in some cases)
    options.add_argument("--no-sandbox")  # Bypass OS security model
    options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource issues in containers

    # Specify the path to chromedriver if needed
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def format_date(date_str):
    # Parse the date string to a datetime object
    if date_str:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        return date_obj.strftime('%B %Y')
    else:
        return 'Present'
    
def remove_specialchars(str):
    if str:
        safe_str = re.sub(r'[<>:"/\\|?*]', '_', str)  # Replace invalid characters
        return safe_str
    else:
        return str