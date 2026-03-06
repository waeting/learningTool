"""
inspect_elements.py
===================
登入 MOOC 平台後，抓取「我修的課」頁面的分頁元件 HTML 供分析。
"""
import getpass, os, sys, time
import mooc_auto
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

os.makedirs("debug", exist_ok=True)

# ── 啟動 headless Chrome ──────────────────────────────────────────────────────
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

# ── 登入 ──────────────────────────────────────────────────────────────────────
username = input("帳號：").strip()
password = getpass.getpass("密碼：")
wait = WebDriverWait(driver, 10)
mooc_auto.start_login(driver, wait, method="教育雲端")
mooc_auto.auto_fill_oauth_form(driver, wait, username, password)

captcha_value = mooc_auto.extract_captcha_and_prompt(driver)
if captcha_value:
    mooc_auto.fill_captcha_and_submit(driver, captcha_value)
    print("[Inspect] 已送出登入，等待導回…")
    time.sleep(8)

if not mooc_auto.verify_login(driver, WebDriverWait(driver, 10)):
    print("[Inspect] 登入失敗。")
    driver.save_screenshot("debug/inspect_login_fail.png")
    driver.quit()
    sys.exit(1)

print("[Inspect] 登入成功，前往我修的課…")

# ── 階段 1：UI 導航進入（模擬 Pass 1）────────────────────────────────────────
wait2 = WebDriverWait(driver, 15)
mooc_auto.click_user_avatar(driver, wait2)
try:
    btn = wait2.until(EC.element_to_be_clickable((By.XPATH,
        "//button[normalize-space(.)='我修的課'] | //a[normalize-space(.)='我修的課']")))
    btn.click()
except Exception:
    driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')

mooc_auto._wait_for_course_list(driver)
mooc_auto._apply_in_progress_filter(driver)

driver.save_screenshot("debug/inspect_phase1.png")
rows1 = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__accordion-head')]")
unpassed_btns1 = driver.find_elements(By.XPATH, "//button[contains(@class,'ml-table__button--unpassed')]")
print(f"[Phase1] header rows: {len(rows1)}, unpassed buttons: {len(unpassed_btns1)}")
for i, r in enumerate(rows1[:3]):
    print(f"  row[{i}] unpassed={mooc_auto._is_row_unpassed(r)}, title={mooc_auto._row_title(r)!r}")

# ── 階段 2：driver.get() 重新載入（模擬 _reload_my_learning）────────────────
print("\n[Inspect] 模擬 _reload_my_learning...")
driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
mooc_auto._wait_for_course_list(driver)
mooc_auto._apply_in_progress_filter(driver)

driver.save_screenshot("debug/inspect_phase2.png")
rows2 = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__accordion-head')]")
unpassed_btns2 = driver.find_elements(By.XPATH, "//button[contains(@class,'ml-table__button--unpassed')]")
print(f"[Phase2] header rows: {len(rows2)}, unpassed buttons: {len(unpassed_btns2)}")
for i, r in enumerate(rows2[:3]):
    print(f"  row[{i}] unpassed={mooc_auto._is_row_unpassed(r)}, title={mooc_auto._row_title(r)!r}")

# 存整頁 HTML 供比對
with open("debug/inspect_phase2_full.html", "w", encoding="utf-8") as f:
    f.write(driver.page_source)
print("[Inspect] Phase2 完整 HTML → debug/inspect_phase2_full.html")

# ── 抓 detail row（accordion 展開內容）HTML ───────────────────────────────────
detail_rows = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__accordion-head')]/following-sibling::tr[1]")
print(f"[Inspect] detail rows in DOM: {len(detail_rows)}")
if detail_rows:
    with open("debug/inspect_detail_row.html", "w", encoding="utf-8") as f:
        f.write(detail_rows[0].get_attribute('outerHTML') or '')
    print("[Inspect] 第一個 detail row HTML → debug/inspect_detail_row.html")

driver.quit()
print("[Inspect] 完成。")
