from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import random
import pandas as pd

def setup_driver():
    chrome_options = Options()
    # Uncomment the line below to run in background (headless mode)
    # chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Disable automation flags
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def scrape_amazon_selenium(keyword: str, domain: str = "ae", pages: int = 1):
    driver = setup_driver()
    all_products = []
    
    try:
        for page in range(1, pages + 1):
            # Build URL for pagination (page=2,3,...)
            base_url = get_amazon_search_url(keyword, domain)
            if page > 1:
                url = f"{base_url}&page={page}"
            else:
                url = base_url
                
            driver.get(url)
            # Random delay to appear human
            time.sleep(random.uniform(3, 5))
            
            # Wait for results to load (Selenium will wait up to 10 seconds)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-component-type='s-search-result']"))
            )
            
            # Find product cards
            products = driver.find_elements(By.CSS_SELECTOR, "[data-component-type='s-search-result']")
            
            for product in products:
                try:
                    # Name
                    name_elem = product.find_element(By.CSS_SELECTOR, "h2 a span")
                    name = name_elem.text
                    
                    # Price (Selenium version)
                    price = "N/A"
                    try:
                        # Check for whole price (a-offscreen is common in search results)
                        price_elem = product.find_element(By.CSS_SELECTOR, ".a-price .a-offscreen")
                        price = price_elem.get_attribute("innerHTML")
                    except:
                        # Try the split price method
                        whole = product.find_elements(By.CSS_SELECTOR, ".a-price-whole")
                        fraction = product.find_elements(By.CSS_SELECTOR, ".a-price-fraction")
                        if whole and fraction:
                            price = f"{whole[0].text}.{fraction[0].text}"
                            
                    all_products.append({"Name": name, "Price": price, "Page": page})
                    print(f"Found: {name} - {price}")
                except Exception as e:
                    print(f"Error parsing card: {e}")
                    continue
            
            # Random delay between pages
            time.sleep(random.uniform(5, 10))
            
    finally:
        driver.quit()
        
    return all_products

# Run the scraper for Amazon.ae
data = scrape_amazon_selenium("WD Purple", "ae", pages=2)
df = pd.DataFrame(data)
print(df)
