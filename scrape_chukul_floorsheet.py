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


def get_first_cell_text(driver):
    # Used to confirm page actually changed after clicking next
    try:
        el = driver.find_element(By.CSS_SELECTOR, "table tbody tr td")
        return el.text.strip()
    except Exception:
        return None


def click_next_page(driver, wait):
    """
    More reliable than clicking page numbers.
    Clicks the 'Next page' arrow until it becomes disabled.
    """
    # Find next button (Quasar pagination)
    # Try exact aria-label first, then fallback contains 'Next'
    try:
        next_btn = driver.find_element(
            By.CSS_SELECTOR, "div.q-pagination__content button[aria-label='Next page']"
        )
    except Exception:
        next_btn = driver.find_element(
            By.CSS_SELECTOR, "div.q-pagination__content button[aria-label*='Next']"
        )

    # If disabled -> last page reached
    if next_btn.get_attribute("disabled") is not None:
        return False

    before = get_first_cell_text(driver)

    driver.execute_script("arguments[0].click();", next_btn)

    # Wait until table content changes (prevents scraping same page again)
    if before is not None:
        wait.until(lambda d: get_first_cell_text(d) != before)
    else:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

    return True


def main():
    # Nepal time for filename
    npt = timezone(timedelta(hours=5, minutes=45))
    run_date = datetime.now(npt).strftime("%Y-%m-%d")

    os.makedirs("outputs", exist_ok=True)
    out_xlsx = f"outputs/floorsheet_{run_date}.xlsx"

    chrome_options = webdriver.ChromeOptions()

    # NOTE: You said you don't want headless on desktop.
    # Keep headless OFF for local run. For GitHub Actions you can turn it ON.
    # chrome_options.add_argument("--headless=new")

    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=chrome_options)  # Selenium Manager auto driver
    wait = WebDriverWait(driver, 30)

    try:
        driver.get(URL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

        all_data = []
        page_count = 0

        while True:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))
            page_data = scrape_current_page(driver)
            all_data.extend(page_data)
            page_count += 1
            print(f"Scraped page: {page_count} | Rows collected (raw): {len(all_data)}")

            # Go next; if not possible => last page done
            if not click_next_page(driver, wait):
                print("Reached last page (Next disabled).")
                break

        df = pd.DataFrame(all_data)
        header = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]

        if df.shape[1] != len(header):
            raise ValueError(f"Column mismatch: got {df.shape[1]} cols, expected {len(header)}")

        df.columns = header
        df["Quantity"] = df["Quantity"].apply(parse_numeric)
        df["Rate"] = df["Rate"].apply(parse_numeric)
        df["Amount"] = df["Amount"].apply(parse_numeric)

        total_amount = df["Amount"].dropna().sum()
        print(f"Pages scraped: {page_count} | Rows: {len(df)} | Total Amount: {total_amount:,.2f}")

        df.to_excel(out_xlsx, index=False)
        print(f"Saved: {out_xlsx}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
