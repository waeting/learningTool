"""
inspect_elements.py
===================
登入 MOOC 平台後，抓取以下元素的 HTML 供分析：
1. CAPTCHA 區域（換一張按鈕）
2. 我修的課頁面的篩選下拉選單
"""
import json, os, sys, time
import mooc_auto
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

os.makedirs("debug", exist_ok=True)

config = mooc_auto.load_config()
username = config.get('username', '')
password = config.get('password', '')
login_method = config.get('login_method', '教育雲端')

options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--window-size=1920,1080')
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--ignore-certificate-errors')

driver = webdriver.Chrome(options=options)
driver.get('https://moocs.moe.edu.tw/moocs/#/home')

mooc_auto.ensure_chinese_language(driver)
time.sleep(5)

wait = WebDriverWait(driver, 10)
mooc_auto.start_login(driver, wait, method=login_method)

if username and password:
    mooc_auto.auto_fill_oauth_form(driver, wait, username, password)

# ── 截圖 + 抓 CAPTCHA 區域 HTML ──────────────────────────────────────────────
time.sleep(1)
driver.save_screenshot("debug/inspect_captcha.png")

# 抓 CAPTCHA 附近的完整 HTML（包含換一張按鈕）
captcha_html_parts = []
for by, sel in [
    (By.XPATH, "//*[contains(translate(@src,'CAPTCHA','captcha'),'captcha') or contains(translate(@src,'CAPTCHA','captcha'),'CheckCode') or contains(translate(@src,'CAPTCHA','captcha'),'validcode') or contains(translate(@src,'CAPTCHA','captcha'),'verify')]"),
    (By.XPATH, "//img[contains(@src,'captcha') or contains(@src,'CheckCode') or contains(@src,'validcode')]"),
]:
    try:
        elems = driver.find_elements(by, sel)
        for el in elems:
            # Walk up to a container that likely holds both the image and the refresh button
            for _ in range(5):
                html = el.get_attribute('outerHTML')
                parent_html = el.find_element(By.XPATH, '..').get_attribute('outerHTML')
                captcha_html_parts.append(f"<!-- captcha img ancestor -->\n{parent_html[:4000]}")
                el = el.find_element(By.XPATH, '..')
    except Exception:
        pass

# Also dump the full form area
try:
    forms = driver.find_elements(By.TAG_NAME, 'form')
    for i, f in enumerate(forms):
        captcha_html_parts.append(f"<!-- form[{i}] -->\n{f.get_attribute('outerHTML')[:6000]}")
except Exception:
    pass

with open("debug/inspect_captcha_area.html", "w", encoding="utf-8") as f:
    f.write("\n\n".join(captcha_html_parts) if captcha_html_parts else "<!-- nothing found -->")
print(f"[Inspect] CAPTCHA 區域 HTML 已儲存 ({len(captcha_html_parts)} 段)")

# Extract captcha to show user
captcha_value = mooc_auto.extract_captcha_and_prompt(driver)
if captcha_value:
    mooc_auto.fill_captcha_and_submit(driver, captcha_value)
    print("[Inspect] 已送出登入表單，等待導回…")
    time.sleep(8)

if not mooc_auto.verify_login(driver, WebDriverWait(driver, 10)):
    print("[Inspect] 登入失敗，無法繼續抓取課程頁元素。")
    driver.save_screenshot("debug/inspect_login_fail.png")
    driver.quit()
    sys.exit(1)

print("[Inspect] 登入成功，前往我修的課…")

# ── 前往我修的課，抓篩選器 HTML ───────────────────────────────────────────────
wait2 = WebDriverWait(driver, 15)
mooc_auto.click_user_avatar(driver, wait2)
try:
    btn = wait2.until(EC.element_to_be_clickable((By.XPATH,
        "//button[normalize-space(.)='我修的課'] | //a[normalize-space(.)='我修的課']")))
    btn.click()
except Exception:
    driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')

time.sleep(5)
driver.save_screenshot("debug/inspect_my_learning.png")

# 抓篩選下拉選單相關 HTML
filter_parts = []
for by, sel in [
    (By.XPATH, "//*[contains(@class,'select') or contains(@class,'filter') or contains(@class,'dropdown')]"),
    (By.TAG_NAME, "select"),
    (By.TAG_NAME, "mat-select"),
    (By.XPATH, "//mat-form-field"),
]:
    try:
        elems = driver.find_elements(by, sel)
        for el in elems[:5]:
            filter_parts.append(f"<!-- {sel} -->\n{el.get_attribute('outerHTML')[:3000]}")
    except Exception:
        pass

with open("debug/inspect_filter.html", "w", encoding="utf-8") as f:
    f.write("\n\n".join(filter_parts) if filter_parts else "<!-- nothing found -->")
print(f"[Inspect] 篩選器 HTML 已儲存 ({len(filter_parts)} 段)")

# 也存整頁 HTML
with open("debug/inspect_my_learning_full.html", "w", encoding="utf-8") as f:
    f.write(driver.page_source)
print("[Inspect] 完整頁面 HTML 已儲存至 debug/inspect_my_learning_full.html")

driver.quit()
print("[Inspect] 完成。")
