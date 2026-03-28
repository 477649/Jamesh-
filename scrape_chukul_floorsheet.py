import os
import re
import time
from datetime import datetime, timezone, timedelta

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://chukul.com/floorsheet"

TABLE_SELECTOR = (
    "#q-app > div > div > div.q-page-container.mobile-padding > "
    "div.q-pull-to-refresh > div.q-pull-to-refresh__content > div > div > div > "
    "div.shadow-inner.full-width.q-px-xs > div > div > table"
)

EXPECTED_HEADER = ["Transact No.", "Symbol", "Buyer", "Seller", "Quantity", "Rate", "Amount"]

# Nepal time and today's date auto-pick
DATE_TO_PICK = "25/03/2026"

def parse_numeric(value):
    """Convert values like '5.27 K' or '2.63 Lac.' into numbers."""
    if value is None:
        return None

    s = str(value).strip().replace(",", "")
    if not s:
        return None

    match = re.search(r"(\d+(?:\.\d+)?)", s)
    if not match:
        return None

    num = float(match.group(1))
    s_lower = s.lower()

    if "lac" in s_lower:
        return num * 100000
    if "k" in s_lower:
        return num * 1000

    return num


def build_driver():
    is_github = os.getenv("GITHUB_ACTIONS") == "true"

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    if is_github:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
    else:
        chrome_options.add_argument("--start-maximized")

    return webdriver.Chrome(options=chrome_options)


def wait_for_table(driver, wait):
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, TABLE_SELECTOR)))
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, f"{TABLE_SELECTOR} tbody tr")))


def get_page_marker(driver):
    """
    Use first row text as page-change marker.
    JS-based to reduce stale element issues.
    """
    script = f"""
    const table = document.querySelector({TABLE_SELECTOR!r});
    if (!table) return null;

    const firstRow = table.querySelector("tbody tr");
    if (!firstRow) return null;

    return firstRow.innerText.trim();
    """
    try:
        return driver.execute_script(script)
    except Exception:
        return None


def scrape_current_page(driver, max_retries=5):
    """
    Fast page scrape using JavaScript.
    Avoids slow td-by-td Selenium reads.
    """
    script = f"""
    const table = document.querySelector({TABLE_SELECTOR!r});
    if (!table) return [];

    const rows = Array.from(table.querySelectorAll("tbody tr"));
    return rows.map(row =>
        Array.from(row.querySelectorAll("td")).map(td => td.innerText.trim())
    );
    """

    for _ in range(max_retries):
        try:
            data = driver.execute_script(script)
            data = [row for row in data if row and len(row) == len(EXPECTED_HEADER)]
            if data:
                return data
        except Exception:
            pass

        time.sleep(0.2)

    return []


def find_next_button(driver):
    """
    Find the real clickable Next button.
    """
    xpath_candidates = [
        "//i[normalize-space()='keyboard_arrow_right']/ancestor::button[1]",
        "//*[@aria-label='Next page']",
    ]

    for xp in xpath_candidates:
        elements = driver.find_elements(By.XPATH, xp)
        if elements:
            return elements[0]

    return None


def click_next_page(driver, wait, max_retries=3):
    """
    Click the Next page button and wait for table content to change.
    Returns False when no more pages are available.
    """
    for _ in range(max_retries):
        next_btn = find_next_button(driver)
        if not next_btn:
            print("Next button not found.")
            return False

        try:
            disabled = next_btn.get_attribute("disabled")
            aria_disabled = next_btn.get_attribute("aria-disabled")
            classes = (next_btn.get_attribute("class") or "").lower()

            if disabled is not None or aria_disabled == "true" or "disabled" in classes:
                print("Next button is disabled. Reached last page.")
                return False

            before = get_page_marker(driver)

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
            time.sleep(0.1)

            try:
                driver.execute_script("arguments[0].click();", next_btn)
            except Exception:
                next_btn.click()

            wait.until(lambda d: get_page_marker(d) not in (None, before))
            return True

        except (TimeoutException, StaleElementReferenceException, ElementClickInterceptedException):
            time.sleep(0.3)

    print("Failed to move to next page after retries.")
    return False


def has_valid_rows(rows):
    for row in rows:
        row_text = " ".join(row).strip().lower()
        if row_text and "no data" not in row_text and "no records" not in row_text:
            return True
    return False


def pick_today_date_and_confirm_data(driver, wait):
    """
    FIRST CODE ONLY:
    Just confirm whether data exists or not for today's date.
    If yes -> return True
    If no -> return False
    """
    print("Opening website for confirmation check...")
    driver.get(URL)
    time.sleep(3)

    day_to_pick = str(int(DATE_TO_PICK.split("/")[0]))
    print("Using date:", DATE_TO_PICK)
    print("Picking day:", day_to_pick)

    calendar_icon = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//i[contains(@class,'material-icons') and normalize-space()='calendar_today']"
        ))
    )
    driver.execute_script("arguments[0].click();", calendar_icon)
    time.sleep(2)

    day_button = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//div[contains(@class,'q-date__calendar-days')]//button[.//span[contains(@class,'block') and normalize-space()='{day_to_pick}']]"
        ))
    )
    driver.execute_script("arguments[0].click();", day_button)
    time.sleep(2)

    try:
        wait_for_table(driver, wait)
        rows = scrape_current_page(driver)

        valid_rows = []
        for row in rows:
            row_text = " ".join(row).strip().lower()
            if row_text and "no data" not in row_text and "no records" not in row_text:
                valid_rows.append(row)

        if valid_rows:
            print("RESULT: yes")
            return True

        print("RESULT: no")
        return False

    except TimeoutException:
        print("RESULT: no")
        return False


def main():
    # Repo root
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # Nepal time
    run_date = datetime.now(npt).strftime("%Y-%m-%d")

    # Output path
    out_dir = os.path.join(base_dir, "outputs", "Floor Sheet")
    os.makedirs(out_dir, exist_ok=True)

    out_csv = os.path.join(out_dir, f"floorsheet_{run_date}.csv")

    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # FIRST CODE: only confirm yes/no
        if not pick_today_date_and_confirm_data(driver, wait):
            print("Stopping script because no data exists.")
            return

        print("Data exists. Running scraper...")

        # SECOND CODE: scrape data
        all_data = []
        page_no = 1

        while True:
            page_data = scrape_current_page(driver)

            if not page_data:
                print(f"No rows found on page {page_no}. Stopping.")
                break

            all_data.extend(page_data)
            print(f"Scraped page {page_no} | rows: {len(page_data)} | total rows: {len(all_data)}")

            moved = click_next_page(driver, wait)
            if not moved:
                print("No more pages found.")
                break

            page_no += 1

        if not all_data:
            raise ValueError("No data scraped from the website.")

        df = pd.DataFrame(all_data, columns=EXPECTED_HEADER)

        df["Quantity"] = df["Quantity"].apply(parse_numeric)
        df["Rate"] = df["Rate"].apply(parse_numeric)

        # Recalculate Amount from Quantity * Rate
        df["Amount"] = df["Quantity"] * df["Rate"]

        df.to_csv(out_csv, index=False, encoding="utf-8-sig")
        print(f"Saved successfully: {out_csv}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
