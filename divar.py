

import time
import uuid
import json
import pandas as pd
from datetime import datetime, timezone
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException


CATEGORY_URL = "https://divar.ir/s/iran/furniture-wood"
CATEGORY_NAME = "مبلمان و صنایع چوب"
SOURCE_NAME = "divar.ir"
CRAWL_DATETIME = datetime.now(timezone.utc).isoformat()
MAX_ADS = 500


def fa_to_en(text):

    fa_digits = '۰۱۲۳۴۵۶۷۸۹'
    en_digits = '0123456789'
    translation_table = str.maketrans(fa_digits, en_digits)
    return text.translate(translation_table)


def get_text_safe(driver, selector, by=By.CSS_SELECTOR):
    try:
        return driver.find_element(by, selector).text.strip()
    except:
        return "N/A"


chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

driver = webdriver.Chrome(options=chrome_options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
})

try:
    print(f"در حال جمع‌آوری لینک‌ها (هدف: {MAX_ADS})...")
    driver.get(CATEGORY_URL)
    time.sleep(4)

    ad_links = set()
    last_height = driver.execute_script("return document.body.scrollHeight")

    while len(ad_links) < MAX_ADS:
        cards = driver.find_elements(By.CSS_SELECTOR, "a[href*='/v/']")
        for card in cards:
            href = card.get_attribute("href")
            if href and "/v/" in href:
                ad_links.add(href)
                if len(ad_links) >= MAX_ADS:
                    break

        if len(ad_links) >= MAX_ADS:
            break

        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        print(f"لینک‌های یافت شده: {len(ad_links)}")

    data_list = []
    final_links = list(ad_links)[:MAX_ADS]

    for idx, url in enumerate(final_links):
        print(f"({idx+1}/{MAX_ADS}) استخراج: {url}")
        driver.get(url)
        time.sleep(2.5)

        record = {
            "doc_id": str(uuid.uuid4()),
            "source": SOURCE_NAME,
            "category": CATEGORY_NAME,
            "city": "N/A",
            "title": "N/A",
            "description": "N/A",
            "price": "N/A",
            "currency": "IRR",
            "post_datetime": "N/A",
            "image_count": 0,
            "url": url,
            "crawl_datetime": CRAWL_DATETIME
        }

        try:
            record["title"] = get_text_safe(driver, "h1")

            subtitle = get_text_safe(driver, ".kt-page-title__subtitle")
            if "در" in subtitle:
                parts = subtitle.split("در")
                record["post_datetime"] = parts[0].strip()
                record["city"] = parts[1].strip()
            else:
                record["city"] = subtitle

            rows = driver.find_elements(By.CLASS_NAME, "kt-base-row")
            for row in rows:
                try:
                    title_text = row.find_element(
                        By.CLASS_NAME, "kt-base-row__title").text
                    if "قیمت" in title_text:
                        value_el = row.find_element(
                            By.CSS_SELECTOR, ".kt-unexpandable-row__value, .kt-base-row__value")
                        record["price"] = value_el.text.strip()
                        break
                except:
                    continue

            try:
                record["description"] = driver.find_element(
                    By.CSS_SELECTOR, "p.kt-description-row__text--primary").text.strip()
            except:
                record["description"] = get_text_safe(
                    driver, ".kt-description-row__text")

            try:
                badge_elements = driver.find_elements(
                    By.CSS_SELECTOR, ".kt-tag__text.kt-text-truncate")
                found_count = False

                for el in badge_elements:
                    text = el.text.strip()
                    if text:
                        if "از" in text:
                            total = text.split("از")[-1].strip()
                            record["image_count"] = int(fa_to_en(total))
                            found_count = True
                            break
                        elif text.isdigit() or all(c in '۰۱۲۳۴۵۶۷۸۹' for c in text):
                            record["image_count"] = int(fa_to_en(text))
                            found_count = True
                            break

                if not found_count:
                    dots = driver.find_elements(
                        By.CSS_SELECTOR, ".kt-carousel__pagination .kt-carousel__dot")
                    thumbs = driver.find_elements(
                        By.CSS_SELECTOR, ".kt-carousel__thumbnails .kt-accent-thumbnail")
                    if len(dots) > 0:
                        record["image_count"] = len(dots)
                    elif len(thumbs) > 0:
                        record["image_count"] = len(thumbs)
                    else:
                        single_img = driver.find_elements(
                            By.CSS_SELECTOR, ".kt-image-block img, .kt-carousel__slide img")
                        record["image_count"] = 1 if len(single_img) > 0 else 0
            except:
                record["image_count"] = 0

        except Exception as e:
            print(f"Error on {url}: {e}")

        data_list.append(record)

    if data_list:
        cols = ["doc_id", "source", "category", "city", "title", "description",
                "price", "currency", "post_datetime", "image_count", "url", "crawl_datetime"]

        df = pd.DataFrame(data_list, columns=cols)
        df.to_excel("dataset.xlsx", index=False)

        with open("dataset.json", "w", encoding="utf-8") as f:
            json.dump(data_list, f, ensure_ascii=False, indent=2)

        print(
            f"\nعملیات با موفقیت پایان یافت. {len(data_list)} آگهی ذخیره شد.")
    else:
        print("داده‌ای یافت نشد.")

finally:
    driver.quit()
