from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
import time

driver = webdriver.Edge(
    service=Service(
        EdgeChromiumDriverManager().install()
    )
)

all_data = []

# Change 50 to 100 or more if needed
for page in range(100, 200):

    url = f"https://www.1mg.com/drugs-all-medicines?page={page}"

    print(f"\nProcessing Page {page}")

    driver.get(url)

    time.sleep(5)

    links = driver.find_elements(By.TAG_NAME, "a")

    page_count = 0

    for link in links:
        try:
            href = link.get_attribute("href")
            text = link.text.strip()

            if href and "/drugs/" in href:

                all_data.append({
                    "Medicine Name": text,
                    "Link": href
                })

                page_count += 1

        except:
            pass

    print(f"Found {page_count} medicines")

driver.quit()

df = pd.DataFrame(all_data)

# Remove duplicates
df = df.drop_duplicates(subset=["Link"])

df.to_excel(
    "all_medicine_urls.xlsx",
    index=False
)

print("\nDone")
print("Saved to all_medicine_urls.xlsx")
print("Total Medicines:", len(df))
