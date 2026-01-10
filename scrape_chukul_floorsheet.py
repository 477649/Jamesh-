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


def go_to_next_page(driver, current_page):
    buttons = driver.find_elements(By.CSS_SELECTOR, "div.q-pagination__middle button")
    target = str(current_page + 1)
    for b in buttons:
        if b.get_attribute("aria-label") == target:
            driver.execute_script("arguments[0].click();", b)
            return True
    return False


def main():
    # Nepal time for filename
    npt = timezone(timedelta(hours=5, minutes=45))
    run_date = datetime.now(npt).strftime("%Y-%m-%d")

    os.makedirs("outputs", exist_ok=True)
    out_xlsx = f"outputs/floorsheet_{run_date}.xlsx"

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

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

            if not go_to_next_page(driver, current_page):
                break

            current_page += 1
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

        df = pd.DataFrame(all_data)
        header = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]

        if df.shape[1] != len(header):
            raise ValueError(f"Column mismatch: got {df.shape[1]} cols, expected {len(header)}")

        df.columns = header
        df["Quantity"] = df["Quantity"].apply(parse_numeric)
        df["Rate"] = df["Rate"].apply(parse_numeric)
        df["Amount"] = df["Amount"].apply(parse_numeric)

        total_amount = df["Amount"].dropna().sum()
        print(f"Rows: {len(df)} | Total Amount: {total_amount:,.2f}")
        df.to_excel(out_xlsx, index=False)
        print(f"Saved: {out_xlsx}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
