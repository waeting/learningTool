"""
debug_run.py
============

Monkey-patch wrapper for mooc_auto.py — Phase 3 onwards.

Replaces key functions with instrumented versions that:
  - Log each strategy attempt individually (pass / fail + reason)
  - Save a screenshot to debug/ at critical checkpoints
  - Dump nav/header outerHTML to debug/ when selectors fail

mooc_auto.py is NOT modified.

Usage:
    python3 debug_run.py [same flags as mooc_auto.py]
    python3 debug_run.py --headless
    python3 debug_run.py --child-headless

Output:
    debug/NN_<label>.png       screenshots
    debug/NN_<label>_nav.html  nav/header HTML dumps
"""

import builtins
import os
import time
import mooc_auto
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

os.makedirs("debug", exist_ok=True)

_TRIGGER_FILE = "debug/login_done.txt"


# ── Replace input() with a file-based trigger ────────────────────────────────
# When running in background (no stdin), input() raises EOFError.
# Instead, we poll for the existence of debug/login_done.txt every second.
# The user creates that file from a second terminal when login is complete.

_orig_input = builtins.input


def _file_trigger_input(prompt=""):
    mooc_auto.log("[Debug] ════════════════════════════════════")
    mooc_auto.log("[Debug] 登入完成後，請在另一個終端機執行：")
    mooc_auto.log(f"[Debug]   touch {_TRIGGER_FILE}")
    mooc_auto.log("[Debug] 或直接建立那個空白檔案")
    mooc_auto.log("[Debug] ════════════════════════════════════")
    while not os.path.exists(_TRIGGER_FILE):
        time.sleep(1)
    try:
        os.remove(_TRIGGER_FILE)
    except Exception:
        pass
    mooc_auto.log("[Debug] 觸發檔案偵測到，繼續執行…")
    return ""


builtins.input = _file_trigger_input

# Sequential counter so files sort in the order they were created
_step = [0]


# ── Snapshot helper ──────────────────────────────────────────────────────────

def _snap(driver: webdriver.Chrome, label: str) -> None:
    """Save a screenshot and nav/header HTML to the debug/ folder."""
    _step[0] += 1
    prefix = f"debug/{_step[0]:02d}_{label}"

    try:
        driver.save_screenshot(f"{prefix}.png")
        mooc_auto.log(f"[Debug] 截圖 → {prefix}.png")
    except Exception as e:
        mooc_auto.log(f"[Debug] 截圖失敗：{e}")

    # Dump the first nav / header / mat-toolbar found on the page
    html_parts = []
    for tag in ("nav", "header", "mat-toolbar"):
        try:
            elems = driver.find_elements(By.TAG_NAME, tag)
            for el in elems:
                html_parts.append(f"<!-- <{tag}> -->\n{el.get_attribute('outerHTML')}")
        except Exception:
            pass

    if html_parts:
        path = f"{prefix}_nav.html"
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n\n".join(html_parts))
            mooc_auto.log(f"[Debug] nav HTML → {path}")
        except Exception as e:
            mooc_auto.log(f"[Debug] HTML 存檔失敗：{e}")
    else:
        mooc_auto.log("[Debug] 頁面上找不到 nav / header / mat-toolbar 元素")


# ── Phase 3: click_user_avatar ───────────────────────────────────────────────
# Rewrite with per-strategy logging so we know exactly which one
# succeeded or why each one failed.

_MY_COURSES_XPATH = (
    "//button[normalize-space(.)='我修的課']"
    " | //a[normalize-space(.)='我修的課']"
)

_AVATAR_STRATEGIES = [
    (
        "1. double-asterisk nav/header button",
        By.XPATH,
        "//nav//button[.//*[contains(text(),'**') or contains(text(),'＊＊')]]"
        " | //header//button[.//*[contains(text(),'**') or contains(text(),'＊＊')]]",
    ),
    (
        "2. single ＊ descendant in nav/header",
        By.XPATH,
        "//*[contains(@class,'nav') or contains(@class,'header')]"
        "//*[contains(text(),'＊')]",
    ),
    (
        "3. person / account_circle mat-icon in nav",
        By.XPATH,
        "//nav//button[.//mat-icon["
        "contains(text(),'person') or contains(text(),'account_circle')]]",
    ),
    (
        "4. element with 'avatar' in class",
        By.XPATH,
        "//*[contains(@class,'avatar')]",
    ),
    (
        "5. nav/header button with 'user' in class",
        By.CSS_SELECTOR,
        "nav button[class*='user'], header button[class*='user']",
    ),
]


def _debug_click_user_avatar(driver: webdriver.Chrome, wait: WebDriverWait) -> bool:
    mooc_auto.log("[Debug] ════ click_user_avatar 開始 ════")
    _snap(driver, "phase3_before_avatar")

    for name, by, selector in _AVATAR_STRATEGIES:
        mooc_auto.log(f"[Debug] 嘗試 {name}")
        try:
            elem = wait.until(EC.element_to_be_clickable((by, selector)))
            mooc_auto.log(
                f"[Debug]   ✓ 元素找到：tag={elem.tag_name!r}"
                f"  class={elem.get_attribute('class')!r:.60}"
                f"  text={elem.text.strip()[:40]!r}"
            )
            elem.click()
            mooc_auto.log("[Debug]   已點擊，等待「我修的課」出現…")
            wait.until(EC.element_to_be_clickable((By.XPATH, _MY_COURSES_XPATH)))
            mooc_auto.log(f"[Debug]   ✓ {name} 成功，下拉選單已開啟")
            _snap(driver, "phase3_dropdown_open")
            return True
        except Exception as e:
            mooc_auto.log(f"[Debug]   ✗ 失敗：{str(e)[:120]}")

    mooc_auto.log("[Debug] ✗ 所有 strategy 失敗，存 nav HTML 供分析")
    _snap(driver, "phase3_all_strategies_failed")
    return False


mooc_auto.click_user_avatar = _debug_click_user_avatar


# ── Phase 4: open_in_progress_courses_mod ───────────────────────────────────
# Wrap to log how many courses were found and snapshot the state after.

_orig_open = mooc_auto.open_in_progress_courses_mod


def _debug_open(driver: webdriver.Chrome, child_headless: bool = False):
    mooc_auto.log("[Debug] ════ open_in_progress_courses_mod 開始 ════")
    result = _orig_open(driver, child_headless=child_headless)
    course_pairs, course_ids = result
    mooc_auto.log(
        f"[Debug] 找到課程數：{len(course_ids)}"
        f"  IDs：{course_ids}"
    )
    _snap(driver, "phase4_after_course_scan")
    return result


mooc_auto.open_in_progress_courses_mod = _debug_open


# ── Phase 5: run_click_loop — 10-second interval ─────────────────────────────

_orig_loop = mooc_auto.run_click_loop


def _debug_loop(triples, interval_seconds=30, stop_event=None):
    mooc_auto.log(
        f"[Debug] run_click_loop 間隔縮短為 10 秒（原 {interval_seconds} 秒）"
    )
    return _orig_loop(triples, interval_seconds=10, stop_event=stop_event)


mooc_auto.run_click_loop = _debug_loop


# ── Phase 5: _toggle_course_tabs ─────────────────────────────────────────────
# Wrap to screenshot on exception (normal path already has [Loop] logs).

_orig_toggle = mooc_auto._toggle_course_tabs


def _debug_toggle(driver: webdriver.Chrome, handle: str, course_id: str) -> bool:
    try:
        return _orig_toggle(driver, handle, course_id)
    except Exception as e:
        mooc_auto.log(f"[Debug] _toggle_course_tabs 例外 ({course_id})：{e}")
        try:
            _snap(driver, f"phase5_toggle_error")
        except Exception:
            pass
        raise


mooc_auto._toggle_course_tabs = _debug_toggle


# ── Phase 5: _check_reading_progress ─────────────────────────────────────────
# Wrap to snapshot the page whenever the function returns True (completion
# detected) so we have visual proof.

_orig_check = mooc_auto._check_reading_progress


def _debug_check(driver: webdriver.Chrome, course_id: str) -> bool:
    result = _orig_check(driver, course_id)
    if result:
        mooc_auto.log(f"[Debug] ✓ 完成條件達成 ({course_id})，存截圖")
        try:
            _snap(driver, f"phase5_course_complete")
        except Exception:
            pass
    return result


mooc_auto._check_reading_progress = _debug_check


# ── Login: wrap start_login to snapshot dialog state ─────────────────────────

def _debug_start_login(driver: webdriver.Chrome, wait, method: str = "教育雲端") -> bool:
    """Debug version: click login, sleep 3 s, snapshot, then select method."""
    mooc_auto.log("[Debug] ════ start_login 開始 ════")

    # Step 1: Click the "登入" nav button
    try:
        login_btn = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//button[normalize-space(.)='登入'"
            " and contains(@class,'action__button--blue')]",
        )))
        login_btn.click()
        mooc_auto.log("[Debug] 已點擊「登入」按鈕。")
    except Exception as e:
        mooc_auto.log(f"[Debug] 找不到「登入」按鈕：{e}")
        return False

    # Snapshot immediately (< 0.5 s) to catch the dialog while it's open
    time.sleep(0.5)
    _snap(driver, "login_immediately_after_click")
    mooc_auto.log(f"[Debug] 視窗數：{len(driver.window_handles)}  handles={driver.window_handles}")

    # Sleep 3 s, then snapshot again to see dialog state
    time.sleep(2.5)
    _snap(driver, "login_after_3s_sleep")
    # Also dump the full page source for DOM inspection
    try:
        path = "debug/login_dialog_3s.html"
        with open(path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        mooc_auto.log(f"[Debug] 3s 後頁面 HTML → {path}")
    except Exception as e:
        mooc_auto.log(f"[Debug] HTML 存檔失敗：{e}")

    # Step 3: Try to click the dialog option
    _METHOD_XPATH = {
        "教育雲端": (
            "//*[contains(normalize-space(.),'教育雲端帳號或縣市帳號')"
            " and not(contains(normalize-space(.),'一般帳號登入'))"
            " and not(contains(normalize-space(.),'TANetRoaming'))]"
        ),
        "一般帳號": (
            "//*[contains(normalize-space(.),'一般帳號登入')"
            " and not(contains(normalize-space(.),'教育雲端帳號或縣市帳號'))"
            " and not(contains(normalize-space(.),'TANetRoaming'))]"
        ),
        "TANetRoaming": (
            "//*[contains(normalize-space(.),'TANetRoaming')"
            " and not(contains(normalize-space(.),'教育雲端帳號或縣市帳號'))"
            " and not(contains(normalize-space(.),'一般帳號登入'))]"
        ),
    }
    xpath = _METHOD_XPATH.get(method, _METHOD_XPATH["教育雲端"])

    for attempt in range(3):
        try:
            elems = driver.find_elements(By.XPATH, xpath)
            mooc_auto.log(f"[Debug] XPath 找到 {len(elems)} 個元素：")
            for i, el in enumerate(elems[:5]):
                mooc_auto.log(
                    f"[Debug]   [{i}] tag={el.tag_name!r}"
                    f"  class={el.get_attribute('class')!r:.60}"
                    f"  text={el.text.strip()[:60]!r}"
                )
            if elems:
                elems[0].click()
                mooc_auto.log(f"[Debug] ✓ 已點擊選項（第 0 個元素），attempt={attempt}")
                _snap(driver, "login_method_clicked")
                return True
            else:
                mooc_auto.log("[Debug] ✗ XPath 找不到任何元素")
                return False
        except Exception as e:
            err = str(e)
            mooc_auto.log(f"[Debug] 點擊失敗（attempt {attempt}）：{err[:120]}")
            if "stale" in err.lower() and attempt < 2:
                time.sleep(0.5)
                continue
            break

    _snap(driver, "login_method_failed")
    try:
        path = "debug/login_dialog_failed.html"
        with open(path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        mooc_auto.log(f"[Debug] 失敗頁面 HTML → {path}")
    except Exception:
        pass
    return False


mooc_auto.start_login = _debug_start_login


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mooc_auto.main()
