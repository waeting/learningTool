"""
mooc_auto_windows.py
=====================

This script is an adaptation of the original MOOC automation script for
the Taiwan Ministry of Education MOOCs platform (磨課師).  It automates
interaction with the platform using Selenium and adheres to the
following workflow:

* The script always opens a **visible** (GUI) Chrome window so the user
  can complete login, including the mandatory CAPTCHA.  After login is
  confirmed, the script can optionally transfer the session to a headless
  Chrome driver if ``--headless`` is supplied.

* After login, the script navigates to the "我修的課" (My Learning) page
  by clicking the user avatar to reveal the dropdown, then clicking the
  "我修的課" link.

* Once on the My Learning page, it locates all courses whose status is
  "進行中" (in progress).  For each such course, the script performs
  the following steps **sequentially**:

    1. Click the course link in the current window.  This triggers
       the Angular router to navigate to a URL of the form
       ``#/learning/<course_id>``.
    2. Capture the full course URL from the address bar and extract
       the course ID.
    3. Open a **new browser window** (not merely a new tab) and load
       the captured course URL in that window.  Each new window
       operates independently and will later be managed by its own
       background thread.
    4. Switch back to the original (parent) window and navigate back
       to the My Learning list to continue processing the next course.

  The script remembers each course ID and the order in which courses
  are opened.  It avoids processing the same course twice by keeping
  track of previously seen course IDs.

* After all in‑progress courses have been opened in separate windows,
  the main thread spawns a background thread for each course window.
  Each thread periodically toggles between the "通過標準" and "課程簡介"
  tabs within its respective course page every 25 minutes.  This
  activity prevents the platform from timing out and ensures that
  reading time is recorded.

**Note:**  The behaviour of the MOOC site may change over time.  The
script uses Chinese text labels (e.g. ``我修的課``, ``進行中``,
``通過標準``, ``課程簡介``) to locate elements.  If these labels
change, you may need to update the locators accordingly.  Also, this
script does **not** attempt to automate login; the user must log in
manually when prompted.
"""

import getpass
import os
import subprocess
import sys
import tempfile
import threading
import time
import re
import argparse
from typing import List, Tuple

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def log(msg: str) -> None:
    """Print a timestamped log message."""
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def _prompt_or_file(prompt: str, trigger_file: str) -> str:
    """Return user input via terminal (interactive) or file trigger (background).

    When stdin is a real TTY, uses input() directly.  Otherwise instructs the
    user to write the value to *trigger_file* from another terminal, then
    reads and removes that file.
    """
    if sys.stdin.isatty():
        return input(prompt)
    log("════════════════════════════════════")
    log(f"請在另一個終端機執行（將你的值填入引號內）：")
    log(f"  echo '你的輸入' > {trigger_file}")
    log("════════════════════════════════════")
    while not os.path.exists(trigger_file):
        time.sleep(1)
    try:
        with open(trigger_file, 'r', encoding='utf-8') as f:
            value = f.read().strip()
        os.remove(trigger_file)
    except Exception:
        value = ""
    log("已讀取輸入，繼續執行…")
    return value


def _prompt_password() -> str:
    """Prompt for a password without echoing it to the terminal.

    In interactive (TTY) mode, uses getpass.getpass() so the input is hidden.
    In background (non-TTY) mode, falls back to the file-trigger pattern:
    the user writes the password to debug/password_input.txt from another
    terminal, and this function reads and immediately deletes that file.
    """
    if sys.stdin.isatty():
        return getpass.getpass("[Login] 請輸入密碼（輸入不會顯示）：")
    os.makedirs("debug", exist_ok=True)
    trigger = "debug/password_input.txt"
    log("════════════════════════════════════")
    log("請在另一個終端機執行（輸入後檔案會立即刪除）：")
    log(f"  echo '你的密碼' > {trigger}")
    log("════════════════════════════════════")
    while not os.path.exists(trigger):
        time.sleep(1)
    try:
        with open(trigger, 'r', encoding='utf-8') as f:
            value = f.read().strip()
        os.remove(trigger)
    except Exception:
        value = ""
    log("已讀取密碼，繼續執行…")
    return value


def ensure_chinese_language(driver: webdriver.Chrome) -> None:
    """Ensure the MOOC platform is displayed in Traditional Chinese.

    If the page is already in Chinese (detected by the presence of "登入"
    link text or the Chinese search placeholder), this function returns
    immediately.  Otherwise it attempts to click the globe/language button
    in the top navigation bar, waits for the dropdown to appear, selects
    "繁體中文(中)", and waits for the page to reload in Chinese.
    """
    wait = WebDriverWait(driver, 10)

    # Check if already in Chinese.
    # "登入" is a <button> in Angular Material, not an <a>, so LINK_TEXT won't
    # find it.  Use XPath text matching instead.
    try:
        driver.find_element(
            By.XPATH,
            "//button[normalize-space(.)='登入']"
            " | //a[normalize-space(.)='登入']",
        )
        return  # Already Chinese
    except Exception:
        pass
    try:
        driver.find_element(By.XPATH, "//*[contains(text(), '您想學習什麼課程')]")
        return
    except Exception:
        pass

    # Try multiple strategies to find the language/globe button.
    # Key distinction from the nav HTML (captured while logged in):
    #   Globe  button: action__button + mat-icon-button, NO action__button--blue
    #   Login  button: action__button + mat-button     + action__button--blue
    #   Bell   button: action__button + mat-icon-button + mat-badge
    #   User   button: action__button + mat-button     + action__button--blue + menu__button
    # So globe = action__button AND mat-icon-button AND NOT mat-badge AND NOT action__button--blue
    globe_strategies = [
        # 1. Most specific: icon-button with action__button, no badge, no blue variant
        (By.XPATH,
         "//nav//button[contains(@class,'action__button')"
         " and contains(@class,'mat-icon-button')"
         " and not(contains(@class,'mat-badge'))"
         " and not(contains(@class,'action__button--blue'))]"),
        # 2. SVG icon id fallback — the globe SVG uses id containing "ic_globe"
        (By.XPATH,
         "//nav//button[.//*[contains(@id,'ic_globe')]]"),
        # 3. aria-label fallback
        (By.XPATH,
         "//button[contains(@aria-label,'language') or contains(@aria-label,'Language')]"),
    ]

    clicked_globe = False
    for by, selector in globe_strategies:
        try:
            btn = wait.until(EC.element_to_be_clickable((by, selector)))
            btn.click()
            clicked_globe = True
            break
        except Exception:
            continue

    if not clicked_globe:
        log("[Lang] 無法找到語言切換按鈕，跳過語言切換。")
        return

    # Wait for dropdown and click Traditional Chinese option
    try:
        zh_option = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'繁體中文')]"))
        )
        zh_option.click()
        # Wait for the page to reload in Chinese
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//button[normalize-space(.)='登入'] | //a[normalize-space(.)='登入']",
            ))
        )
        log("[Lang] 語言已切換為繁體中文。")
    except Exception as e:
        log(f"[Lang] 切換繁體中文失敗：{e}")


def verify_login(driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
    """Return True if the user appears to be logged in.

    Checks for the disappearance of the "登入" button (which is replaced
    by the user's name after login) or the presence of a full-width asterisk
    character (＊) which appears in masking username formats like "王＊＊".
    """
    try:
        wait.until(EC.invisibility_of_element_located((
            By.XPATH,
            "//button[normalize-space(.)='登入']",
        )))
        return True
    except Exception:
        pass
    try:
        driver.find_element(By.XPATH, "//*[contains(text(),'＊')]")
        return True
    except Exception:
        return False


def start_login(driver: webdriver.Chrome, wait: WebDriverWait,
                method: str = "教育雲端") -> bool:
    """Click the Login button and select a login method from the dialog.

    Clicks the "登入" button in the nav bar, waits for the method-selection
    dialog to appear, then clicks the option matching *method*.  The user
    still needs to fill in their credentials and CAPTCHA manually.

    Supported values for *method*:
        "教育雲端"   – 使用教育雲端帳號或縣市帳號登入 (default)
        "一般帳號"   – 使用教育雲端一般帳號登入
        "TANetRoaming" – 使用臺灣學術網路無線漫遊登入

    Returns True if the method was selected successfully, False otherwise.
    """
    _METHOD_XPATH = {
        # The login option rows are <a class="login-nav__provider-link"> elements.
        # Prefer the <a> tag selector; fall back to text-exclusion if DOM changes.
        "教育雲端": (
            "//a[contains(@class,'login-nav__provider-link')"
            " and contains(normalize-space(.),'教育雲端帳號或縣市帳號')]"
            " | //*[contains(normalize-space(.),'教育雲端帳號或縣市帳號')"
            " and not(contains(normalize-space(.),'一般帳號登入'))"
            " and not(contains(normalize-space(.),'TANetRoaming'))]"
        ),
        "一般帳號": (
            "//a[contains(@class,'login-nav__provider-link')"
            " and contains(normalize-space(.),'一般帳號登入')]"
            " | //*[contains(normalize-space(.),'一般帳號登入')"
            " and not(contains(normalize-space(.),'教育雲端帳號或縣市帳號'))"
            " and not(contains(normalize-space(.),'TANetRoaming'))]"
        ),
        "TANetRoaming": (
            "//a[contains(@class,'login-nav__provider-link')"
            " and contains(normalize-space(.),'TANetRoaming')]"
            " | //*[contains(normalize-space(.),'TANetRoaming')"
            " and not(contains(normalize-space(.),'教育雲端帳號或縣市帳號'))"
            " and not(contains(normalize-space(.),'一般帳號登入'))]"
        ),
    }

    # Click the nav-bar "登入" button
    try:
        login_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[normalize-space(.)='登入'"
            " and contains(@class,'action__button--blue')]",
        )))
        login_btn.click()
        log("[Login] 已點擊「登入」按鈕。")
    except Exception as e:
        log(f"[Login] 找不到「登入」按鈕：{e}")
        return False

    # Wait for the method-selection dialog and click the chosen option.
    # Angular Material dialogs animate in; wait for animation to fully
    # settle before attempting to click, then retry on stale element.
    time.sleep(3)
    xpath = _METHOD_XPATH.get(method, _METHOD_XPATH["教育雲端"])
    for attempt in range(3):
        try:
            method_elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            method_elem.click()
            log(f"[Login] 已選擇登入方式：{method}。")
            return True
        except Exception as e:
            err = str(e)
            if "stale" in err.lower() and attempt < 2:
                time.sleep(0.5)
                continue
            log(f"[Login] 找不到登入方式「{method}」：{err[:120]}")
            return False
    return False


def click_user_avatar(driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
    """Click the user avatar/name to expand the dropdown containing "我修的課".

    Based on the observed page structure: the trigger is a <button> in the
    nav bar containing a person mat-icon, the masked username (e.g. "楊**"),
    and a chevron icon.  The dropdown items are Angular Material
    <button mat-menu-item> elements, NOT <a> tags, so LINK_TEXT cannot be
    used to detect or click them.

    Tries multiple selectors in order.  Returns True if the dropdown was
    successfully opened (i.e. "我修的課" button became clickable).
    """
    # XPath that matches "我修的課" whether it is a <button> or an <a>.
    # normalize-space(.) collects all descendant text, so it works even
    # when the label is inside a nested <span>.
    my_courses_xpath = (
        "//button[normalize-space(.)='我修的課'] | //a[normalize-space(.)='我修的課']"
    )

    strategies = [
        # 1. Nav/header button whose descendant text contains ** (e.g. "楊**")
        #    — matches both half-width ** and full-width ＊＊
        (By.XPATH,
         "//nav//button[.//*[contains(text(),'**') or contains(text(),'＊＊')]]"
         " | //header//button[.//*[contains(text(),'**') or contains(text(),'＊＊')]]"),
        # 2. Any nav/header descendant containing a full-width asterisk ＊
        #    (catches ＊＊ as well since it contains ＊)
        (By.XPATH,
         "//*[contains(@class,'nav') or contains(@class,'header')]"
         "//*[contains(text(),'＊')]"),
        # 3. Nav button with a "person" or "account_circle" mat-icon child
        #    — matches the person-icon button seen in the screenshot
        (By.XPATH,
         "//nav//button[.//mat-icon[contains(text(),'person')"
         " or contains(text(),'account_circle')]]"),
        # 4. Any element whose class contains "avatar"
        (By.XPATH, "//*[contains(@class,'avatar')]"),
        # 5. Nav/header button with a "user"-related class
        (By.CSS_SELECTOR,
         "nav button[class*='user'], header button[class*='user']"),
    ]
    for by, selector in strategies:
        try:
            elem = wait.until(EC.element_to_be_clickable((by, selector)))
            elem.click()
            # Confirm the dropdown opened by waiting for the "我修的課" item.
            # Use XPath (not LINK_TEXT) because the item is a <button>, not <a>.
            wait.until(EC.element_to_be_clickable((By.XPATH, my_courses_xpath)))
            return True
        except Exception:
            continue
    return False



def auto_fill_oauth_form(driver: webdriver.Chrome, wait: WebDriverWait,
                         username: str, password: str) -> bool:
    """Auto-fill credentials on the OAuth provider page after redirect.

    Tries common field selectors for username/password inputs.
    Returns True if both fields were filled successfully.
    """
    log("[Login] 等待 OAuth 頁面載入…")
    time.sleep(2)

    # --- Username ---
    # Also try switching into iframes in case the form is embedded
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    contexts_to_try = [None] + frames  # None = top-level; frames = each iframe

    username_elem = None
    for frame in contexts_to_try:
        try:
            if frame is not None:
                driver.switch_to.frame(frame)
            else:
                driver.switch_to.default_content()
        except Exception:
            continue
        for by, sel in [
            (By.XPATH, "//input[@placeholder='請輸入帳號']"),
            (By.CSS_SELECTOR, "input[type='text']:not([disabled]):not([readonly])"),
            (By.CSS_SELECTOR, "input[type='email']:not([disabled]):not([readonly])"),
            (By.CSS_SELECTOR, "input:not([type]):not([disabled]):not([readonly])"),
            (By.XPATH, "//input[not(@type='password') and not(@type='hidden') and not(@disabled) and not(@readonly)]"),
        ]:
            try:
                username_elem = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((by, sel)))
                break
            except Exception:
                continue
        if username_elem:
            break

    if not username_elem:
        driver.switch_to.default_content()
        log("[Login] 找不到帳號欄位，跳過自動填入。")
        return False

    try:
        username_elem.clear()
        username_elem.send_keys(username)
        log("[Login] 已填入帳號。")
    except Exception as e:
        log(f"[Login] 填入帳號失敗：{e}")
        return False

    # --- Password ---
    try:
        pwd_elem = driver.find_element(By.XPATH, "//input[@type='password']")
        pwd_elem.clear()
        pwd_elem.send_keys(password)
        log("[Login] 已填入密碼。")
    except Exception as e:
        log(f"[Login] 找不到或填入密碼失敗：{e}")
        return False

    return True


def _find_captcha_img(driver: webdriver.Chrome):
    """Return the CAPTCHA <img> element, or None if not found."""
    for by, sel in [
        (By.ID, "id2b"),  # Confirmed id from DOM inspection
        (By.XPATH, "//img[contains(translate(@src,'CAPTCHA','captcha'),'captcha')]"),
        (By.XPATH, "//img[contains(@src,'CheckCode') or contains(@src,'validcode')]"),
        (By.XPATH, "//img[contains(translate(@src,'CAPTCHA','captcha'),'verify')]"),
    ]:
        try:
            return driver.find_element(by, sel)
        except Exception:
            continue
    return None


def extract_captcha_and_prompt(driver: webdriver.Chrome) -> str:
    """Find the CAPTCHA image, save it to debug/captcha.png, open it for the
    user, then return the CAPTCHA string typed by the user in the terminal.

    If the user enters 'r' (or writes 'r' to the trigger file), the function
    clicks the '換下一個' button on the login page, waits for the CAPTCHA
    image to refresh, and prompts again.  This loop repeats until the user
    provides a non-'r' value.
    """
    os.makedirs("debug", exist_ok=True)

    while True:
        captcha_img = _find_captcha_img(driver)
        if captcha_img:
            try:
                captcha_img.screenshot('debug/captcha.png')
                log("[Login] CAPTCHA 圖片已儲存至 debug/captcha.png，正在開啟預覽…")
                subprocess.Popen(['open', 'debug/captcha.png'])
            except Exception as e:
                log(f"[Login] 儲存 CAPTCHA 圖片失敗：{e}")
        else:
            log("[Login] 未找到 CAPTCHA 圖片，請直接查看瀏覽器。")

        log("[Login] 輸入驗證碼，或輸入 r 換一張：")
        value = _prompt_or_file("[Login] 驗證碼（r=換一張）：", "debug/captcha_input.txt")

        if value.strip().lower() != 'r':
            return value.strip()

        # User wants a new CAPTCHA — click '換下一個'
        log("[Login] 正在切換下一張 CAPTCHA…")
        old_src = captcha_img.get_attribute('src') if captcha_img else None
        try:
            refresh_btn = driver.find_element(By.ID, "id12")
            refresh_btn.click()
        except Exception:
            try:
                refresh_btn = driver.find_element(
                    By.XPATH, "//a[normalize-space(.)='換下一個' or @title='換下一個']"
                )
                refresh_btn.click()
            except Exception as e:
                log(f"[Login] 找不到換下一個按鈕：{e}")
                continue

        # Step 1: wait for the img src to change (antiCache param updates)
        if old_src:
            try:
                WebDriverWait(driver, 5).until(
                    lambda d: (
                        _find_captcha_img(d) is not None
                        and _find_captcha_img(d).get_attribute('src') != old_src
                    )
                )
            except Exception:
                time.sleep(1)
        else:
            time.sleep(1)

        # Step 2: wait for the new image to finish loading (naturalWidth > 0)
        new_img = _find_captcha_img(driver)
        if new_img:
            try:
                WebDriverWait(driver, 5).until(
                    lambda d: driver.execute_script(
                        "var el = arguments[0];"
                        "return el.complete && el.naturalWidth > 0;",
                        new_img,
                    )
                )
            except Exception:
                time.sleep(1)  # fallback sleep if JS check fails
        log("[Login] CAPTCHA 已刷新。")


def fill_captcha_and_submit(driver: webdriver.Chrome, captcha_value: str) -> bool:
    """Fill the CAPTCHA input and submit the login form.

    Returns True if submission appeared to succeed.
    """
    # Find CAPTCHA input field
    captcha_input = None
    for by, sel in [
        (By.XPATH, "//input[contains(translate(@name,'CAPTCHA','captcha'),'captcha') or contains(translate(@id,'CAPTCHA','captcha'),'captcha')]"),
        (By.XPATH, "//input[contains(translate(@placeholder,'CAPTCHA驗證碼','captcha驗證碼'),'驗證碼') or contains(translate(@placeholder,'CAPTCHA驗證碼','captcha驗證碼'),'captcha')]"),
        (By.XPATH, "//input[contains(@name,'CheckCode') or contains(@name,'ValidCode') or contains(@name,'verifyCode')]"),
        (By.XPATH, "(//input[@type='text'])[last()]"),  # fallback: last text input on page
    ]:
        try:
            captcha_input = driver.find_element(by, sel)
            break
        except Exception:
            continue

    if not captcha_input:
        log("[Login] 找不到驗證碼欄位，無法自動填入。")
        return False

    try:
        captcha_input.clear()
        captcha_input.send_keys(captcha_value)
        log("[Login] 已填入驗證碼。")
    except Exception as e:
        log(f"[Login] 填入驗證碼失敗：{e}")
        return False

    # Submit the form
    for by, sel in [
        (By.XPATH, "//button[@type='submit']"),
        (By.XPATH, "//input[@type='submit']"),
        (By.XPATH, "//button[normalize-space(.)='登入' or normalize-space(.)='確認' or normalize-space(.)='Submit' or normalize-space(.)='Login']"),
    ]:
        try:
            btn = driver.find_element(by, sel)
            btn.click()
            log("[Login] 已點擊送出按鈕。")
            return True
        except Exception:
            continue

    # Last resort: submit via the input element
    try:
        captcha_input.submit()
        log("[Login] 已送出表單（submit）。")
        return True
    except Exception as e:
        log(f"[Login] 送出表單失敗：{e}")
        return False


def transfer_to_headless_via_profile(user_data_dir: str) -> webdriver.Chrome:
    """Open a new headless Chrome driver reusing an existing profile directory.

    The caller must have already called gui_driver.quit() before calling this
    function so that Chrome has released its lock on the profile directory.

    Args:
        user_data_dir: Path to the Chrome user-data directory from the previous
            GUI session.  The directory must still exist (not auto-deleted).

    Returns:
        A new headless Chrome WebDriver using the same profile.
    """
    log(f"[Main] 以 headless 模式重新開啟 Chrome（使用 profile：{user_data_dir}）…")
    headless_options = webdriver.ChromeOptions()
    headless_options.add_argument(f'--user-data-dir={user_data_dir}')
    headless_options.add_argument('--headless=new')
    headless_options.add_argument('--window-size=1920,1080')
    headless_options.add_argument('--no-sandbox')
    headless_options.add_argument('--disable-gpu')
    headless_options.add_argument('--disable-dev-shm-usage')
    headless_options.add_argument('--ignore-certificate-errors')
    # Must match the GUI driver's password-store setting so encrypted cookies
    # are stored with the same key and can be read in headless mode.
    headless_options.add_argument('--password-store=basic')
    headless_options.add_argument('--no-first-run')
    headless_options.add_argument('--no-default-browser-check')
    driver = webdriver.Chrome(options=headless_options)
    driver.get('https://moocs.moe.edu.tw/moocs/#/home')
    # Wait for Angular SPA to initialise and authentication state to restore
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//nav | //app-navbar"))
        )
    except Exception:
        pass
    time.sleep(5)  # Extra buffer for Angular to finish rendering the nav
    return driver


# ---------------------------------------------------------------------------
# Core automation
# ---------------------------------------------------------------------------

# Track the minute count at which each course first reached 100%.
# Key: course_id  Value: minute count when 100% was first seen
_first_100_pct_minutes: dict = {}


def _debug_progress_elements(driver: webdriver.Chrome, course_id: str) -> None:
    """Log debug information about progress-related elements on the current page."""
    all_progress = driver.find_elements(
        By.XPATH,
        "//*[contains(@class,'progress') or contains(text(),'閱讀') or contains(text(),'分鐘')]",
    )
    if all_progress:
        log(f"[Loop] 找到 {len(all_progress)} 個可能相關的元素 ({course_id})")
        for i, elem in enumerate(all_progress[:3]):
            log(f"[Loop] 元素 {i+1}: {elem.tag_name} - {elem.text[:50]} ({course_id})")
    else:
        log(f"[Loop] 頁面上沒有找到任何相關元素 ({course_id})")


def _check_reading_progress(driver: webdriver.Chrome, course_id: str) -> bool:
    """Read the 閱讀時數 progress on the currently visible 通過標準 tab.

    Assumes the driver is already switched to the correct window and the
    通過標準 tab is active.

    Returns:
        True if the course has reached 100 % completion AND the recorded
        minute count is divisible by 5 (the platform's safe-to-close
        condition).  False in all other cases.
    """
    try:
        progress_elements = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'course-status__progress')]//span[contains(@class,'course-status__progress-label') and contains(text(),'閱讀時數')]/following-sibling::*//div[contains(@class,'course-status__progress-info')]",
        )

        if not progress_elements:
            reading_blocks = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'course-status__progress')]//span[contains(text(),'閱讀時數')]",
            )
            for block in reading_blocks:
                parent = block.find_element(
                    By.XPATH, "ancestor::div[contains(@class,'course-status__progress')][1]"
                )
                info_divs = parent.find_elements(
                    By.XPATH, ".//div[contains(@class,'course-status__progress-info')]"
                )
                if info_divs:
                    progress_elements = info_divs
                    break

        if not progress_elements:
            progress_elements = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'course-status__progress') and .//span[contains(text(),'閱讀時數')]]//div[contains(@class,'course-status__progress-info')]",
            )

        if not progress_elements:
            log(f"[Loop] 所有策略都未找到閱讀時數進度元素 ({course_id})")
            _debug_progress_elements(driver, course_id)
            return False

        progress_info = progress_elements[0]
        log(f"[Loop] 進度區塊內容: '{progress_info.text.strip()}' ({course_id})")

        minutes_elem = progress_info.find_elements(By.XPATH, ".//span[contains(text(),'分鐘')]")
        percentage_elem = progress_info.find_elements(By.XPATH, ".//small[contains(text(),'%')]")
        if not minutes_elem:
            minutes_elem = progress_info.find_elements(By.XPATH, ".//*[contains(text(),'分鐘')]")
        if not percentage_elem:
            percentage_elem = progress_info.find_elements(By.XPATH, ".//*[contains(text(),'%')]")

        if not (minutes_elem and percentage_elem):
            log(f"[Loop] 找到進度元素但無法解析分鐘數或百分比 ({course_id}): {progress_info.text}")
            return False

        minutes_text = minutes_elem[0].text.strip()
        percentage_text = percentage_elem[0].text.strip()
        log(f"[Loop] 閱讀時數進度：{minutes_text} {percentage_text} ({course_id})")

        if "(100%)" not in percentage_text and "100%" not in percentage_text:
            # Not yet complete — reset any stored first-seen record so that
            # if a course somehow goes back below 100% we start fresh.
            _first_100_pct_minutes.pop(course_id, None)
            return False

        minutes_match = re.search(r'(\d+)', minutes_text)
        if not minutes_match:
            log(f"[Loop] 課程已達到100%但無法解析分鐘數，繼續追蹤 ({course_id})")
            return False

        minutes_number = int(minutes_match.group(1))

        if course_id not in _first_100_pct_minutes:
            # First time we see 100% — record the minute count and wait for
            # the platform to tick up at least one more minute to confirm
            # that the required total has genuinely been reached.
            _first_100_pct_minutes[course_id] = minutes_number
            log(f"[Loop] 初見100%（{minutes_number}分鐘），等下一輪確認分鐘數增加 ({course_id})")
            return False

        first_minutes = _first_100_pct_minutes[course_id]
        if minutes_number > first_minutes:
            log(f"[Loop] 確認完成：100% 且分鐘數已從 {first_minutes} 增至 {minutes_number}，關閉分頁 ({course_id})")
            return True

        log(f"[Loop] 100% 但分鐘數({minutes_number})尚未超過初見值({first_minutes})，繼續等待 ({course_id})")
        return False

    except Exception as e:
        log(f"[Loop] 抓取閱讀時數進度時發生錯誤 ({course_id}): {e}")
        return False


def _toggle_course_tabs(driver: webdriver.Chrome, handle: str, course_id: str) -> bool:
    """Toggle the 通過標準 / 課程簡介 tabs for one course window.

    Assumes the driver has already been switched to ``handle`` before this
    call.  Always ends on 通過標準 so that _check_reading_progress() runs
    every cycle:
      - If currently on 通過標準: click 課程簡介 (visit it), wait 2 s,
        click back to 通過標準.
      - If currently on 課程簡介: click 通過標準 directly (already visited
        課程簡介 this round).

    Returns:
        True if _check_reading_progress() reports the course is complete and
        the window should be closed.  False otherwise.
    """
    wait = WebDriverWait(driver, 30)
    pass_selector = (
        By.XPATH,
        "//div[@role='tab'][div[contains(@class,'mat-tab-label-content') and contains(normalize-space(.),'通過標準')]]",
    )
    intro_selector = (
        By.XPATH,
        "//div[@role='tab'][div[contains(@class,'mat-tab-label-content') and contains(normalize-space(.),'課程簡介')]]",
    )

    pass_elem = wait.until(EC.element_to_be_clickable(pass_selector))
    intro_elem = wait.until(EC.element_to_be_clickable(intro_selector))

    pass_active = 'mat-tab-label-active' in pass_elem.get_attribute('class').split()

    if pass_active:
        # Currently on 通過標準: visit 課程簡介 first, then return
        intro_elem.click()
        log(f"[Loop] 已切換到：課程簡介 ({course_id})")
        time.sleep(2)
        pass_elem.click()
        log(f"[Loop] 返回：通過標準 ({course_id})")
    else:
        # Currently on 課程簡介: go to 通過標準 directly
        pass_elem.click()
        log(f"[Loop] 已切換到：通過標準 ({course_id})")

    # Wait for the page to settle, then confirm we are on 通過標準
    time.sleep(5)
    pass_now = driver.find_element(*pass_selector)
    if 'mat-tab-label-active' in pass_now.get_attribute('class').split():
        log(f"[Loop] 現在停在：通過標準 ({course_id})")
        return _check_reading_progress(driver, course_id)
    log(f"[Loop] 停留在通過標準失敗，跳過進度檢查 ({course_id})")
    return False


def run_click_loop(
    triples: List[Tuple[webdriver.Chrome, str, str]],
    interval_seconds: int = 30,
    stop_event: "threading.Event | None" = None,
) -> None:
    """Sequentially monitor all course windows, toggling tabs every interval.

    This single-threaded sequential approach replaces the previous design
    that spawned one thread per window.  The old design caused race
    conditions when multiple threads shared the same driver instance
    one thread would call switch_to.window() while another was mid-operation,
    causing reads from the wrong window.

    Args:
        triples: List of (driver, window_handle, course_id).  Entries are
            removed in-place as courses complete.
        interval_seconds: Seconds to wait after each full pass over all
            active windows.  Uses stop_event.wait() so Ctrl-C interrupts
            the sleep immediately.
        stop_event: Optional threading.Event.  The loop exits early when
            the event is set (e.g. after KeyboardInterrupt in main).
    """
    if stop_event is None:
        stop_event = threading.Event()

    active: List[Tuple[webdriver.Chrome, str, str]] = list(triples)
    # Track consecutive errors per course; remove only after MAX_ERRORS failures.
    error_counts: dict = {}
    MAX_ERRORS = 3

    while active and not stop_event.is_set():
        for entry in list(active):  # iterate over a snapshot so removal is safe
            if stop_event.is_set():
                break
            drv, handle, course_id = entry

            # Verify the window is still open / switchable
            try:
                drv.switch_to.window(handle)
            except Exception as e:
                count = error_counts.get(course_id, 0) + 1
                error_counts[course_id] = count
                if count >= MAX_ERRORS:
                    log(f"[Loop] 視窗 {course_id} 連續 {MAX_ERRORS} 次無法切換，放棄追蹤：{e}")
                    active.remove(entry)
                    error_counts.pop(course_id, None)
                else:
                    log(f"[Loop] 視窗 {course_id} 切換失敗（{count}/{MAX_ERRORS}），下輪重試：{e}")
                continue

            # Toggle tabs and check for completion
            try:
                done = _toggle_course_tabs(drv, handle, course_id)
                error_counts.pop(course_id, None)  # reset on success
            except Exception as e:
                count = error_counts.get(course_id, 0) + 1
                error_counts[course_id] = count
                log(f"[Loop] 處理課程 {course_id} 時發生錯誤（{count}/{MAX_ERRORS}）：{e}")
                done = False
                if count >= MAX_ERRORS:
                    log(f"[Loop] 課程 {course_id} 連續 {MAX_ERRORS} 次錯誤，放棄追蹤。")
                    active.remove(entry)
                    error_counts.pop(course_id, None)
                    continue

            if done:
                try:
                    drv.close()
                except Exception:
                    pass
                active.remove(entry)
                error_counts.pop(course_id, None)
                log(f"[Loop] 課程 {course_id} 已完成並關閉，剩餘 {len(active)} 門課程繼續追蹤。")

        if active and not stop_event.is_set():
            log(f"[Loop] 本輪結束，{interval_seconds} 秒後繼續（剩餘 {len(active)} 門）。")
            stop_event.wait(timeout=interval_seconds)

    log("[Loop] 所有課程皆已完成或監控已停止。")


def _is_row_unpassed(row) -> bool:
    """Return True if the course row (or its sibling detail row) has the unpassed button."""
    try:
        if row.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
            return True
    except Exception:
        pass
    try:
        detail_row = row.find_element(By.XPATH, "following-sibling::tr[1]")
        if detail_row.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
            return True
    except Exception:
        pass
    return False


def _row_title(row) -> str:
    """Extract the course title text from a header row."""
    try:
        return row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]").text.strip()
    except Exception:
        pass
    try:
        return row.find_element(By.XPATH, ".//a").text.strip()
    except Exception:
        return ""


def _apply_in_progress_filter(driver) -> None:
    """Apply the '進行中' filter on the My Learning page and wait for the list to render."""
    # Use structural XPath instead of dynamic ID (Angular reassigns mat-select-N on each load)
    filter_xpath = "//mat-select[not(ancestor::mat-paginator)]"
    try:
        filter_select = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, filter_xpath))
        )
        filter_select.click()
        try:
            in_progress_option = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//mat-option[contains(normalize-space(.),'進行中')]"))
            )
            in_progress_option.click()
            log("[Navigate] 已套用「進行中」篩選。")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'table__accordion-head')]"))
            )
            time.sleep(1)
        except Exception:
            # Close any open overlay before giving up
            try:
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            except Exception:
                pass
            log("[Navigate] 無法選取「進行中」選項，繼續使用全部課程。")
    except Exception as e:
        log(f"[Navigate] 找不到篩選器（繼續使用全部課程）：{e}")


def _wait_for_course_list(driver) -> None:
    """Wait for the course list rows and their unpassed buttons to be fully rendered."""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'table__accordion-head')]"))
        )
    except Exception:
        pass
    # Wait for Angular to finish rendering the unpassed status buttons inside each row.
    # These are in the detail rows and may appear later than the header rows.
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'ml-table__button--unpassed')]"))
        )
    except Exception:
        pass  # No unpassed courses on this page, or buttons rendered elsewhere
    time.sleep(1)


def _collect_page_titles(driver) -> List[str]:
    """Return unpassed course titles from the current page."""
    try:
        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'table__accordion-head')]")
    except Exception:
        return []
    titles = []
    for row in rows:
        if _is_row_unpassed(row):
            t = _row_title(row)
            if t:
                titles.append(t)
    return titles


def _find_row_on_current_page(driver, title: str):
    """Search the current page for a row whose title matches. Returns the row or None."""
    try:
        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'table__accordion-head')]")
    except Exception:
        return None
    for row in rows:
        if _is_row_unpassed(row) and _row_title(row) == title:
            return row
    return None


def _find_row_across_pages(driver, title: str):
    """Search across all paginator pages for a matching unpassed row. Returns the row or None."""
    while True:
        _wait_for_course_list(driver)
        row = _find_row_on_current_page(driver, title)
        if row is not None:
            return row
        try:
            next_btn = driver.find_element(By.XPATH,
                "//button[contains(@class,'mat-paginator-navigation-next')]")
            if next_btn.get_attribute("disabled"):
                return None
            next_btn.click()
        except Exception:
            return None


def _reload_my_learning(driver) -> None:
    """Navigate back to My Learning page and re-apply the 進行中 filter.

    If already on the my-learning page, navigate via home first to force Angular
    to destroy and recreate the component (resetting the paginator to page 1).
    Navigating to the same hash URL from the same page may be treated as a no-op
    by the Angular router, leaving the paginator on a non-first page.
    When coming from a different route (e.g. a course page), a direct navigation
    is sufficient.
    """
    if 'my-learning' in driver.current_url:
        driver.get('https://moocs.moe.edu.tw/moocs/#/home')
        time.sleep(1)
    driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
    _wait_for_course_list(driver)
    _apply_in_progress_filter(driver)


def open_in_progress_courses_mod(driver: webdriver.Chrome) -> Tuple[List[Tuple[webdriver.Chrome, str]], List[str]]:
    """Navigate to the My Learning page and open all not-yet-passed courses.

    Args:
        driver: A Selenium WebDriver already logged into the MOOCs platform.

    Returns:
        A tuple containing (course_window_pairs, course_ids) where each pair
        is (driver_instance, window_handle).
    """
    wait = WebDriverWait(driver, 20)

    my_courses_xpath = (
        "//button[normalize-space(.)='我修的課'] | //a[normalize-space(.)='我修的課']"
    )
    if click_user_avatar(driver, wait):
        try:
            my_courses_link = wait.until(EC.element_to_be_clickable((By.XPATH, my_courses_xpath)))
            my_courses_link.click()
        except Exception as e:
            log(f"[Navigate] 找到頭像下拉選單但無法點擊「我修的課」：{e}")
            driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
    else:
        log("[Navigate] 無法點擊頭像展開選單，嘗試直接導向我修的課頁面。")
        driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')

    _wait_for_course_list(driver)
    log("[Navigate] 課程列表已載入。")
    _apply_in_progress_filter(driver)

    parent_handle = driver.current_window_handle
    course_windows: List[Tuple[webdriver.Chrome, str]] = []
    course_ids: List[str] = []
    seen_ids: set = set()
    titles: List[str] = []

    # Pass 1: collect all unpassed course titles across all pages
    seen_titles: set = set()
    page_num = 1
    while True:
        _wait_for_course_list(driver)
        page_titles = _collect_page_titles(driver)
        new_titles = [t for t in page_titles if t not in seen_titles]
        seen_titles.update(new_titles)
        titles.extend(new_titles)
        log(f"[Navigate] 第 {page_num} 頁找到 {len(page_titles)} 筆，新增 {len(new_titles)} 筆未完成課程。")
        try:
            next_btn = driver.find_element(By.XPATH,
                "//button[contains(@class,'mat-paginator-navigation-next')]")
            if next_btn.get_attribute("disabled"):
                break
            next_btn.click()
            page_num += 1
        except Exception:
            break

    log(f"[Navigate] 共收集 {len(titles)} 筆未完成課程：{titles}")

    # Pass 2a: collect all course URLs (filter applied once; use driver.back() between clicks)
    _reload_my_learning(driver)
    course_url_map: List[Tuple[str, str, str]] = []  # (title, course_id, course_url)
    for title in titles:
        if not title:
            continue
        matching_row = _find_row_across_pages(driver, title)
        if matching_row is None:
            log(f"[Navigate] 找不到課程「{title}」，略過。")
            # Navigate to page 1 without filter so next search can proceed
            driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
            _wait_for_course_list(driver)
            continue
        try:
            try:
                click_elem = matching_row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]")
            except Exception:
                click_elem = matching_row.find_element(By.XPATH, ".//a")
            click_elem.click()
        except Exception:
            driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
            _wait_for_course_list(driver)
            continue
        try:
            WebDriverWait(driver, 15).until(lambda d: "/learning/" in d.current_url)
        except Exception:
            driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
            _wait_for_course_list(driver)
            continue
        course_url = driver.current_url
        m = re.search(r"(\d+)$", course_url)
        course_id = m.group(1) if m else course_url
        if course_id not in seen_ids:
            seen_ids.add(course_id)
            course_url_map.append((title, course_id, course_url))
        # Go back to my-learning via browser history — Angular resets to page 1.
        # No filter re-application needed: _find_row_across_pages searches all pages.
        driver.back()
        _wait_for_course_list(driver)

    log(f"[Navigate] 共確認 {len(course_url_map)} 筆課程 URL，開始開啟分頁…")

    # Pass 2b: open all collected course URLs in new tabs at once
    for title, course_id, course_url in course_url_map:
        try:
            driver.switch_to.new_window('window')
            driver.get(course_url)
            n_handle = driver.current_window_handle
        except Exception:
            driver.execute_script("window.open(arguments[0], '_blank');", course_url)
            n_handle = driver.window_handles[-1]
            driver.switch_to.window(n_handle)
        course_windows.append((driver, n_handle))
        course_ids.append(course_id)
        log(f"[Navigate] 已開啟課程視窗 {n_handle}，課程 ID {course_id}（{title}）。")
        driver.switch_to.window(parent_handle)

    if not course_windows:
        log("[Navigate] 沒有找到標示為未完成的課程。")
    return course_windows, course_ids


def main() -> None:
    """Entry point for the automation script.

    This function parses command‑line arguments, always starts Chrome with a
    visible GUI (so the user can complete login including CAPTCHA), confirms
    login, optionally transfers the session to a headless driver, then
    collects in‑progress courses, opens each in a new window, and starts
    background threads to handle periodic tab toggling.

    --headless flag semantics (new):
        Login always happens in a GUI window.  After login is confirmed, the
        session cookies are transferred to a new headless Chrome driver and
        the GUI window is closed.  The headless driver then continues all
        remaining automation.

    """
    parser = argparse.ArgumentParser(description="Automate interactions with the MOE MOOCs platform")
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run Chrome in headless mode (no visible window) throughout the entire session."
    )
    args = parser.parse_args()

    login_method = "教育雲端"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')

    if args.headless:
        # ── Headless mode ────────────────────────────────────────────────────
        # Prompt credentials before opening the browser so the user can paste
        # them immediately.  CAPTCHA is extracted as an image file and opened
        # in Preview; the user types the code back via terminal.
        log("[Main] 以 headless 模式啟動 Chrome。")
        if sys.stdin.isatty():
            username = input("[Login] 請輸入帳號：").strip()
        else:
            username = _prompt_or_file("[Login] 請輸入帳號：", "debug/username_input.txt")
        password = _prompt_password()

        options.add_argument('--headless=new')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-dev-shm-usage')
        driver = webdriver.Chrome(options=options)
        driver.get('https://moocs.moe.edu.tw/moocs/#/home')

        ensure_chinese_language(driver)

        _MAX_LOGIN_ATTEMPTS = 6
        login_wait = WebDriverWait(driver, 10)
        for _attempt in range(_MAX_LOGIN_ATTEMPTS):
            time.sleep(1)
            if start_login(driver, login_wait, method=login_method):
                break
            log(f"[Main] 登入對話框未成功開啟，重試（{_attempt + 1}/{_MAX_LOGIN_ATTEMPTS}）…")
        else:
            log("[Main] 警告：已達登入重試上限，嘗試繼續執行。")

        auto_fill_oauth_form(driver, login_wait, username, password)
        captcha_value = extract_captcha_and_prompt(driver)
        if captcha_value:
            fill_captcha_and_submit(driver, captcha_value)
            log("[Main] 已送出登入表單，等待 OAuth 導回並載入…")
            time.sleep(8)
        else:
            log("[Main] 未輸入驗證碼，請手動完成登入後通知腳本繼續。")
            _prompt_or_file("[Main] 登入完成後按 Enter：", "debug/login_done.txt")

        if not verify_login(driver, login_wait):
            log("[Main] 未偵測到登入狀態，請確認已成功登入後重新執行。")
            driver.quit()
            return

    else:
        # ── GUI mode ─────────────────────────────────────────────────────────
        # Open a visible browser, navigate to the login page, and poll until
        # the user completes login manually (no credential prompts, no CAPTCHA
        # extraction — the user interacts directly with the browser).
        log("[Main] 以 GUI 模式啟動 Chrome。")
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()
        driver.get('https://moocs.moe.edu.tw/moocs/#/home')

        ensure_chinese_language(driver)

        _MAX_LOGIN_ATTEMPTS = 6
        login_wait = WebDriverWait(driver, 10)
        for _attempt in range(_MAX_LOGIN_ATTEMPTS):
            time.sleep(1)
            if start_login(driver, login_wait, method=login_method):
                break
            log(f"[Main] 登入對話框未成功開啟，重試（{_attempt + 1}/{_MAX_LOGIN_ATTEMPTS}）…")
        else:
            log("[Main] 警告：已達登入重試上限，請手動點擊登入按鈕。")

        log("[Main] 請在瀏覽器中完成登入（帳號、密碼、驗證碼）…")
        poll_wait = WebDriverWait(driver, 5)
        while not verify_login(driver, poll_wait):
            time.sleep(3)
        login_wait = WebDriverWait(driver, 10)

    log("[Main] 登入確認成功。")

    # Capture the initial window handle so we can close course tabs between
    # cycles while keeping the main (navigation) tab alive.
    main_handle = driver.current_window_handle

    # Every CYCLE_MINUTES we close all course tabs and reopen them, to avoid
    # the platform's 30-minute progress-reset mechanism.
    CYCLE_MINUTES = 25

    stop_event = threading.Event()
    monitor_thread: "threading.Thread | None" = None

    try:
        while True:
            course_pairs, course_ids = open_in_progress_courses_mod(driver)
            if not course_pairs:
                log("[Main] 沒有找到未完成課程，所有課程已完成或無法開啟。")
                break

            # Build triples (driver, handle, course_id) for the monitoring loop.
            triples: List[Tuple[webdriver.Chrome, str, str]] = [
                (drv, hdl, cid) for (drv, hdl), cid in zip(course_pairs, course_ids)
            ]

            stop_event = threading.Event()
            monitor_thread = threading.Thread(
                target=run_click_loop,
                args=(triples, 30, stop_event),
                daemon=False,
            )
            monitor_thread.start()
            log(f"[Main] 已開始監控 {len(triples)} 門課程（順序為找到的順序）：{course_ids}")
            log(f"[Main] 將在 {CYCLE_MINUTES} 分鐘後自動重開課程頁面以避免進度重置。按 Ctrl+C 可中止。")

            # Block until all courses complete OR the cycle timer expires.
            monitor_thread.join(timeout=CYCLE_MINUTES * 60)

            if not monitor_thread.is_alive():
                # Monitoring loop exited naturally: all courses done.
                log("[Main] 所有課程已完成，程式結束。")
                break

            # Cycle timer expired: stop the monitor, close course tabs, then reopen.
            log(f"[Main] 已達 {CYCLE_MINUTES} 分鐘，重新整理課程頁面以避免進度重置…")
            stop_event.set()
            monitor_thread.join(timeout=15)

            # Close every tab except the main navigation tab.
            for handle in list(driver.window_handles):
                if handle != main_handle:
                    try:
                        driver.switch_to.window(handle)
                        driver.close()
                    except Exception:
                        pass
            try:
                driver.switch_to.window(main_handle)
            except Exception:
                pass
            log("[Main] 課程分頁已關閉，重新開啟未完成課程…")
            # Loop continues: open_in_progress_courses_mod will re-open only
            # courses that are still "進行中" (completed ones won't appear).

    except KeyboardInterrupt:
        log("[Main] 接收到中斷訊號，通知監控執行緒停止…")
        if monitor_thread is not None and monitor_thread.is_alive():
            stop_event.set()
            monitor_thread.join(timeout=10)
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == '__main__':
    main()
