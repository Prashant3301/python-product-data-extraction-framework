"""
Phase 1 - Product URL Discovery Demo

This demonstration script shows how a URL discovery workflow can be
implemented for catalog-style pages.

IMPORTANT:
- Uses a public demo site instead of a commercial retailer.
- Intended for learning Selenium automation patterns.
- Only automate sites that permit automated access.
"""

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.microsoft import EdgeChromiumDriverManager

import pandas as pd
import time

# ==========================================================
# CONFIGURATION
# ==========================================================

START_PAGE = 1
END_PAGE = 5
OUTPUT_FILE = "product_urls.xlsx"

# Public demonstration catalog site
CATALOG_URL_TEMPLATE = (
    "https://books.toscrape.com/catalogue/page-{}.html"
)

# ==========================================================
# CREATE BROWSER SESSION
# ==========================================================

driver = webdriver.Edge(
    service=Service(
        EdgeChromiumDriverManager().install()
    )
)

all_data = []

# ==========================================================
# DISCOVER PRODUCT URLS
# ==========================================================

for page in range(START_PAGE, END_PAGE + 1):

    url = CATALOG_URL_TEMPLATE.format(page)

    print(f"
Processing Page {page}")

    driver.get(url)

    time.sleep(3)

    links = driver.find_elements(By.TAG_NAME, "a")

    page_count = 0

    for link in links:
        try:
            href = link.get_attribute("href")
            text = link.text.strip()

            # Demo product-page filter
            if href and "/catalogue/" in href:

                all_data.append({
                    "Product Name": text,
                    "Product URL": href,
                })

                page_count += 1

        except Exception:
            pass

    print(f"Found {page_count} product URLs")

# ==========================================================
# CLEANUP
# ==========================================================

driver.quit()

# ==========================================================
# EXPORT RESULTS
# ==========================================================

df = pd.DataFrame(all_data)

df = df.drop_duplicates(subset=["Product URL"])

df.to_excel(
    OUTPUT_FILE,
    index=False
)

print("
Discovery Completed")
print(f"Output File: {OUTPUT_FILE}")
print(f"Total Product URLs: {len(df)}")
