import os
from datetime import datetime, timezone, timedelta
import pandas as pd
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://chukul.com/floorsheet"


def parse_numeric(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    clean = "".join(c for c in s if c.isdigit() or c == ".")
    try:
        return float(clean) if clean else None
    except ValueError:
        return None


def scrape_current_page(driver):
    soup = BeautifulSoup(driver.page_source, "html.parser")
    table = soup.find("table")
    if not table:
        return []

    rows = table.find_all("tr")
    data = []
    for row in rows:
        cols = row.find_all("td")
        cols_data = [c.get_text(strip=True) for c in cols]
        if cols_data:
            data.append(cols_data)
    return data


def first_row_key(driver):
    """Used to confirm page actually changed after clicking next."""
    try:
        return driver.find_element(By.CSS_SELECTOR, "table tbody tr td").text.strip()
    except Exception:
        return None


def go_to_next_page(driver, wait, current_page):
    """
    Click next page by VISIBLE TEXT (e.g., '139'), not aria-label.
    This fixes the issue where it stops at 138 even though 139 exists.
    """
    target = str(current_page + 1)

    # Find the next page button anywhere inside q-pagination by its visible number
    buttons = driver.find_elements(
        By.XPATH,
        f"//div[contains(@class,'q-pagination')]//button[normalize-space()='{target}']"
    )

    if not buttons:
        return False  # no next page visible => last page reached

    before = first_row_key(driver)

    driver.execute_script("arguments[0].click();", buttons[0])

    # Wait until table content changes (prevents scraping same page again)
    if before is not None:
        wait.until(lambda d: first_row_key(d) != before)
    else:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

    return True


def main():
    # Nepal time for filename
    npt = timezone(timedelta(hours=5, minutes=45))
    run_date = datetime.now(npt).strftime("%Y-%m-%d")

    os.makedirs("outputs", exist_ok=True)
    out_xlsx = f"outputs/floorsheet_{run_date}.xlsx"

    # Auto detect GitHub Actions (must run headless there)
    IS_GITHUB = os.getenv("GITHUB_ACTIONS") == "true"

    chrome_options = webdriver.ChromeOptions()
    if IS_GITHUB:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
    else:
        # Desktop: visible browser
        chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=chrome_options)  # Selenium Manager auto driver
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(URL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

        all_data = []
        current_page = 1

        while True:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
            all_data.extend(scrape_current_page(driver))
            print(f"Scraped page: {current_page}")

            if not go_to_next_page(driver, wait, current_page):
                print("Reached last page.")
                break

            current_page += 1

        df = pd.DataFrame(all_data)
        header = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]

        if df.shape[1] != len(header):
            raise ValueError(f"Column mismatch: got {df.shape[1]} cols, expected {len(header)}")

        df.columns = header
        df["Quantity"] = df["Quantity"].apply(parse_numeric)
        df["Rate"] = df["Rate"].apply(parse_numeric)
        df["Amount"] = df["Amount"].apply(parse_numeric)

        total_amount = df["Amount"].dropna().sum()
        print(f"Pages scraped: {current_page} | Rows: {len(df)} | Total Amount: {total_amount:,.2f}")

        df.to_excel(out_xlsx, index=False)
        print(f"Saved: {out_xlsx}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
