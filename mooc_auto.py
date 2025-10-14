"""
mooc_auto_windows.py
=====================

This script is an adaptation of the original MOOC automation script for
the Taiwan Ministry of Education MOOCs platform (磨課師).  It automates
interaction with the platform using Selenium and adheres to the
following workflow:

* After the user manually logs in, it navigates to the "我修的課"
  (My Learning) page.  If the navigation link is hidden behind a
  user‑profile drop‑down or a hamburger menu, the script will attempt
  to reveal it by clicking the user name (usually a surname followed by
  two asterisks, such as "王＊＊") or the hamburger toggle.

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

import threading
import time
import re
import argparse
from typing import List, Tuple

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def click_periodically(
    driver: webdriver.Chrome,
    window_handle: str,
    interval_seconds: int = 90,
) -> None:
    """Periodically toggles between the "通過標準" and "課程簡介" tabs.

    Args:
        driver: The Selenium WebDriver instance controlling the browser.
        window_handle: The window handle of the course window to operate on.
        interval_seconds: Number of seconds to wait between each toggle cycle.
            Defaults to 30 seconds.  Each cycle performs two clicks (one on
            each tab), waits a few seconds for the interface to update, prints
            the current active tab to the console, then sleeps for
            ``interval_seconds`` before repeating.

    In each iteration this function:
      1. Switches to the specified window.
      2. Locates the "通過標準" and "課程簡介" tabs using XPath selectors.
      3. Determines which tab is currently active.  It clicks the other tab
         first, then clicks back to the original tab so that both tabs
         register interaction.
      4. Waits 5 seconds, prints which tab is active, then sleeps for
         ``interval_seconds`` seconds before starting the next cycle.
    """
    wait = WebDriverWait(driver, 30)
    # Angular Material tab selectors: locate the <div role="tab"> whose
    # nested .mat-tab-label-content contains the desired text.
    pass_selector = (
        By.XPATH,
        "//div[@role='tab'][div[contains(@class, 'mat-tab-label-content') and contains(normalize-space(.), '通過標準')]]",
    )
    intro_selector = (
        By.XPATH,
        "//div[@role='tab'][div[contains(@class, 'mat-tab-label-content') and contains(normalize-space(.), '課程簡介')]]",
    )
    while True:
        try:
            driver.switch_to.window(window_handle)
        except Exception as switch_err:
            print(f"[ClickLoop] Failed to switch to window {window_handle}: {switch_err}")
            return
        try:
            pass_elem = wait.until(EC.element_to_be_clickable(pass_selector))
            intro_elem = wait.until(EC.element_to_be_clickable(intro_selector))
            # Check which tab is currently active (mat-tab-label-active on the role="tab" element)
            pass_active = 'mat-tab-label-active' in pass_elem.get_attribute('class').split()
            intro_active = 'mat-tab-label-active' in intro_elem.get_attribute('class').split()
            # Click the inactive tab first, then click the other tab
            if pass_active and not intro_active:
                # Currently on "通過標準": click "課程簡介" first
                intro_elem.click()
                # Immediately report that we switched to 課程簡介
                print(f"[ClickLoop] 已切換到：課程簡介 (Window {window_handle})")
                time.sleep(2)
                # Then click back to "通過標準"
                pass_elem.click()
            elif intro_active and not pass_active:
                # Currently on "課程簡介": click "通過標準" first
                pass_elem.click()
                print(f"[ClickLoop] 已切換到：通過標準 (Window {window_handle})")
                time.sleep(2)
                # Then click back to "課程簡介"
                intro_elem.click()
            else:
                # If neither or both are active (unexpected), click both and report
                pass_elem.click()
                print(f"[ClickLoop] 已切換到：通過標準 (Window {window_handle})")
                time.sleep(2)
                intro_elem.click()
            # Give the page a little time to settle and determine the active tab
            time.sleep(5)
            # Re-read the class attribute to see which tab is active after toggling
            # (we may have returned to the original active tab).  We must fetch
            # the elements again to get updated class lists.
            pass_elem_current = driver.find_element(*pass_selector)
            intro_elem_current = driver.find_element(*intro_selector)
            pass_active_now = 'mat-tab-label-active' in pass_elem_current.get_attribute('class').split()
            intro_active_now = 'mat-tab-label-active' in intro_elem_current.get_attribute('class').split()
            if pass_active_now and not intro_active_now:
                print(f"[ClickLoop] 現在停在：通過標準 (Window {window_handle})")
            elif intro_active_now and not pass_active_now:
                print(f"[ClickLoop] 現在停在：課程簡介 (Window {window_handle})")
            else:
                print(f"[ClickLoop] 無法確定當前頁籤 (Window {window_handle})")
        except Exception as e:
            print(f"[ClickLoop] Error while clicking tabs in window {window_handle}: {e}")
        # Sleep for the specified interval before starting the next cycle
        time.sleep(interval_seconds)


def open_in_progress_courses_old(driver: webdriver.Chrome) -> Tuple[List[str], List[str]]:
    """Navigate to "我修的課" and open all courses marked as "進行中" in new windows.

    Args:
        driver: The Selenium WebDriver instance controlling the browser.

    Returns:
        A tuple containing two lists:
          * A list of window handles corresponding to the new course windows.
          * A list of course IDs in the order they were opened.

    This function assumes that the user is already logged into the
    platform.  It first looks for the "我修的課" link and clicks it.  If
    the link isn’t immediately visible, it attempts to reveal it by
    clicking the user name (containing asterisks) or the hamburger menu.
    Once on the My Learning page, it processes each course with status
    "進行中": clicking the course to load its URL, opening that URL in a
    new window, and then returning to the course list to repeat.  It
    stops when all unique in‑progress courses have been opened.
    """
    wait = WebDriverWait(driver, 20)
    # Attempt to reveal the user menu by clicking the surname+asterisks.
    try:
        user_name_elem = driver.find_element(By.XPATH, "//*[contains(text(), '＊')]")
        user_name_elem.click()
    except Exception:
        try:
            user_name_elem = driver.find_element(By.XPATH, "//*[contains(text(), '*')]")
            user_name_elem.click()
        except Exception:
            pass

    # Attempt to click on "我修的課" directly.  If not visible, try via user menu or hamburger.
    try:
        my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
    except Exception:
        try:
            user_menu = driver.find_element(
                By.XPATH,
                "//*[(contains(text(), '＊') or contains(text(), '*')) and (self::span or self::button or self::a)]"
            )
            user_menu.click()
            my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
        except Exception:
            try:
                hamburger = driver.find_element(By.CSS_SELECTOR, "button.navbar-toggler, button.hamburger, button[aria-label='Toggle navigation']")
                hamburger.click()
                my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
            except Exception as e:
                print(f"[Navigate] Unable to find '我修的課': {e}")
                return [], []
    my_courses_link.click()
    # Allow the course list to render
    time.sleep(3)
    course_windows: List[str] = []
    course_ids: List[str] = []
    # Track which course IDs have been processed to avoid duplicates
    seen_ids = set()
    # Save the handle of the parent window (the My Learning page)
    parent_handle = driver.current_window_handle
    # Loop until no new courses are found
    # Locate all table rows that have a status label containing "進行中".  We
    # iterate through these rows by index, re‑fetching the rows each time
    # because navigation may refresh the DOM.  For each row, we click its
    # course link, capture the course ID/URL, open it in a new window, and
    # return to the list to process the next row.  This approach avoids
    # missing courses due to dynamic DOM updates.
    try:
        # Each course entry appears as a header row with class "table_accordion-head"
        # followed by a container row.  We select only the header rows containing
        # the "進行中" status label.  This reduces the chance of picking up
        # container rows that don't contain the clickable course title.
        initial_rows = driver.find_elements(
            By.XPATH,
            "//tr[contains(@class, 'table_accordion-head') and .//*[contains(text(), '進行中')]]",
        )
    except Exception:
        initial_rows = []
    if not initial_rows:
        print("[Navigate] 沒有找到標示為未完成的課程。")
        return [], []
    # Iterate over indices rather than storing stale WebElement references
    for index in range(len(initial_rows)):
        try:
            # Re‑fetch the list of rows each iteration to avoid stale elements
            rows = driver.find_elements(
                By.XPATH,
                "//tr[contains(@class, 'table_accordion-head') and .//*[contains(text(), '進行中')]]",
            )
            if index >= len(rows):
                break
            row = rows[index]
            # Find the clickable course name within the row.  The course
            # title appears in a <p> element with classes containing
            # "course-name".  There is no <a> anchor in this table row, so
            # clicking the <p> (or its parent <td>) triggers navigation.
            try:
                course_link = row.find_element(
                    By.XPATH,
                    ".//p[contains(@class, 'course-name')]",
                )
            except Exception:
                continue
            # Click the course link in the parent window
            try:
                course_link.click()
            except Exception:
                continue
            # Wait for the URL to change to a learning page
            try:
                WebDriverWait(driver, 15).until(lambda d: "/learning/" in d.current_url)
            except Exception:
                # If navigation fails, go back and continue with next row
                try:
                    driver.back()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
                except Exception:
                    pass
                continue
            course_url = driver.current_url
            # Extract course ID from URL
            match = re.search(r"(\d+)$", course_url)
            course_id = match.group(1) if match else course_url
            # Skip duplicate courses
            if course_id in seen_ids:
                try:
                    driver.back()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
                except Exception:
                    pass
                continue
            seen_ids.add(course_id)
            course_ids.append(course_id)
            # Open the course URL in a new window
            try:
                driver.switch_to.new_window('window')
                driver.get(course_url)
                new_handle = driver.current_window_handle
            except Exception:
                driver.execute_script("window.open(arguments[0], '_blank');", course_url)
                new_handle = driver.window_handles[-1]
                driver.switch_to.window(new_handle)
            course_windows.append(new_handle)
            print(f"[Navigate] Opened course window {new_handle} for course ID {course_id}.")
            # Return to the parent handle and navigate back to the course list
            try:
                driver.switch_to.window(parent_handle)
                driver.back()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
            except Exception:
                driver.get('https://moocs.moe.edu.tw/moocs/#/my-learning')
                time.sleep(3)
        except Exception:
            continue
    return course_windows, course_ids

#
# New implementation for collecting and opening in-progress courses.
# This function uses the updated DOM structure of the MOOCs platform where
# course rows are identified by the class ``table__accordion-head`` and
# in‑progress courses contain a button with the ``ml-table__button--unpassed``
# class.  Each course title is contained in a ``<p>`` element whose class
# includes ``course-name``.  The function logs the titles of all
# in‑progress courses for debugging and opens each course in a new
# browser window, avoiding duplicates.  It returns the window handles
# and course IDs in the order processed.
def open_in_progress_courses_mod(driver: webdriver.Chrome, child_headless: bool = False) -> Tuple[List[Tuple[webdriver.Chrome, str]], List[str]]:
    """Navigate to the My Learning page and open all not-yet-passed courses.

    Args:
        driver: A Selenium WebDriver already logged into the MOOCs platform.

    Returns:
        A tuple containing (course_window_handles, course_ids).
    """
    wait = WebDriverWait(driver, 20)
    # Attempt to reveal the user menu by clicking the surname followed by asterisks.
    try:
        user_elem = driver.find_element(By.XPATH, "//*[contains(text(), '＊')]")
        user_elem.click()
    except Exception:
        try:
            user_elem = driver.find_element(By.XPATH, "//*[contains(text(), '*')]")
            user_elem.click()
        except Exception:
            pass
    # Ensure the "我修的課" link is visible and click it.
    try:
        my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
    except Exception:
        try:
            user_menu = driver.find_element(
                By.XPATH,
                "//*[(contains(text(), '＊') or contains(text(), '*')) and (self::span or self::button or self::a)]",
            )
            user_menu.click()
            my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
        except Exception:
            try:
                hamburger = driver.find_element(
                    By.CSS_SELECTOR,
                    "button.navbar-toggler, button.hamburger, button[aria-label='Toggle navigation']",
                )
                hamburger.click()
                my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
            except Exception as e:
                print(f"[Navigate] Unable to find '我修的課': {e}")
                return [], []
    my_courses_link.click()
    # Allow the page to load
    time.sleep(3)
    parent_handle = driver.current_window_handle
    # We'll return a list of (driver, handle) pairs.  If child_headless is False,
    # the driver will always be the main driver; otherwise a new headless driver
    # will be spawned for each course and added here.
    course_windows: List[Tuple[webdriver.Chrome, str]] = []
    course_ids: List[str] = []
    seen_ids = set()
    # Log course titles for debugging
    titles = []
    try:
        header_rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'table__accordion-head')]")
    except Exception:
        header_rows = []
    for row in header_rows:
        # Determine if this row represents an in‑progress course
        unpassed = False
        try:
            if row.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
                unpassed = True
        except Exception:
            pass
        if not unpassed:
            try:
                detail_row = row.find_element(By.XPATH, "following-sibling::tr[1]")
                if detail_row.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
                    unpassed = True
            except Exception:
                pass
        if not unpassed:
            continue
        try:
            title_elem = row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]")
            t = title_elem.text.strip()
            if t:
                titles.append(t)
        except Exception:
            try:
                title_elem_alt = row.find_element(By.XPATH, ".//a")
                t = title_elem_alt.text.strip()
                if t:
                    titles.append(t)
            except Exception:
                pass
    if titles:
        print(f"[Navigate] 課程列表：{titles}")
    # Iterate through each course title that we detected as unpassed and open its page sequentially.
    # Get cookies from the main driver so they can be reused in child drivers.
    cookies: List[dict] = []
    if child_headless:
        try:
            cookies = driver.get_cookies()
        except Exception:
            cookies = []
    for title in titles:
        if not title:
            continue
        # Iterate over all header rows to find the one whose title matches exactly
        matching_row = None
        try:
            current_rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'table__accordion-head')]")
        except Exception:
            current_rows = []
        for row in current_rows:
            # Check if the row (or its detail row) has the unpassed button
            unpassed = False
            try:
                if row.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
                    unpassed = True
            except Exception:
                pass
            if not unpassed:
                try:
                    drow = row.find_element(By.XPATH, "following-sibling::tr[1]")
                    if drow.find_elements(By.XPATH, ".//button[contains(@class, 'ml-table__button--unpassed')]"):
                        unpassed = True
                except Exception:
                    pass
            if not unpassed:
                continue
            # Retrieve the title text
            try:
                t_elem = row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]")
                t_text = t_elem.text.strip()
            except Exception:
                try:
                    t_elem = row.find_element(By.XPATH, ".//a")
                    t_text = t_elem.text.strip()
                except Exception:
                    continue
            if t_text == title:
                matching_row = row
                break
        if matching_row is None:
            continue
        # Click the course title element in the matching row
        try:
            # Prefer <p class="course-name"> element
            try:
                click_elem = matching_row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]")
            except Exception:
                click_elem = matching_row.find_element(By.XPATH, ".//a")
            click_elem.click()
        except Exception:
            continue
        # Wait until the URL indicates the course page
        try:
            WebDriverWait(driver, 15).until(lambda d: "/learning/" in d.current_url)
        except Exception:
            # If navigation fails, reload the list and move to the next title
            try:
                driver.switch_to.window(parent_handle)
                driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                time.sleep(3)
            except Exception:
                pass
            continue
        course_url = driver.current_url
        m = re.search(r"(\d+)$", course_url)
        course_id = m.group(1) if m else course_url
        if course_id in seen_ids:
            try:
                driver.switch_to.window(parent_handle)
                driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                time.sleep(3)
            except Exception:
                pass
            continue
        seen_ids.add(course_id)
        course_ids.append(course_id)
        # Depending on child_headless flag, open the course in a new headless driver or a new window in the main driver
        if child_headless:
            child_options = webdriver.ChromeOptions()
            child_options.add_argument('--ignore-certificate-errors')
            child_options.add_argument('--headless')
            child_options.add_argument('--window-size=1920,1080')
            # Enhanced stability options for headless mode
            child_options.add_argument('--disable-gpu')
            child_options.add_argument('--no-sandbox')
            child_options.add_argument('--disable-dev-shm-usage')
            child_options.add_argument('--disable-extensions')
            child_options.add_argument('--disable-web-security')
            child_options.add_argument('--remote-debugging-port=0')  # Let system auto-assign port
            try:
                child_driver = webdriver.Chrome(options=child_options)
            except Exception:
                try:
                    driver.switch_to.window(parent_handle)
                    driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                    time.sleep(3)
                except Exception:
                    time.sleep(3)
                continue
            try:
                child_driver.get('https://moocs.moe.edu.tw')
                for c in cookies:
                    try:
                        cookie_dict = {k: c[k] for k in c if k in ['name','value','domain','path','expiry','secure','httpOnly']}
                        child_driver.add_cookie(cookie_dict)
                    except Exception:
                        continue
                child_driver.get(course_url)
            except Exception:
                try:
                    child_driver.quit()
                except Exception:
                    pass
                try:
                    driver.switch_to.window(parent_handle)
                    driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                    time.sleep(3)
                except Exception:
                    time.sleep(3)
                continue
            c_handle = child_driver.current_window_handle
            course_windows.append((child_driver, c_handle))
            print(f"[Navigate] Opened course window {c_handle} for course ID {course_id} (課程名稱：{title}) in headless child.")
            # Return to list page via reload
            try:
                driver.switch_to.window(parent_handle)
                driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                time.sleep(3)
            except Exception:
                time.sleep(3)
            continue
        else:
            try:
                driver.switch_to.new_window('window')
                driver.get(course_url)
                n_handle = driver.current_window_handle
            except Exception:
                driver.execute_script("window.open(arguments[0], '_blank');", course_url)
                n_handle = driver.window_handles[-1]
                driver.switch_to.window(n_handle)
            course_windows.append((driver, n_handle))
            print(f"[Navigate] Opened course window {n_handle} for course ID {course_id} (課程名稱：{title}).")
            try:
                driver.switch_to.window(parent_handle)
                driver.get('https://moocs.moe.edu.tw/moocs/#/course/my-learning')
                time.sleep(3)
            except Exception:
                time.sleep(3)
    if not course_windows:
        print("[Navigate] 沒有找到標示為未完成的課程。")
    return course_windows, course_ids

# Revised implementation using a fresh scan of all course rows.  This new
# function collects every course on the "我修的課" page that still has an
# "未通過" (not passed) status and opens each one in its own window.
def open_in_progress_courses(driver: webdriver.Chrome) -> Tuple[List[str], List[str]]:
    """Navigate to the My Learning page and open all not-yet-passed courses.

    The function first clicks the "我修的課" link, then iteratively scans the
    list of courses.  Each course entry is represented by a header row
    (``table_accordion-head``) and a detail row (``table_accordion-container``).
    Courses that have not yet been passed display a button with class
    ``ml-table_button--unpassed`` in either the header row or its detail row.
    For each such course, this function:

    1. Clicks the course title (found in a ``<p>`` with class including
       ``course-name``) to trigger navigation to a URL of the form
       ``#/learning/<course_id>``.
    2. Extracts the course ID from the URL and checks whether it has
       already been processed.
    3. Opens the course URL in a new browser window (or a new tab if
       necessary) and records the window handle.
    4. Returns to the My Learning page to continue processing the next
       course.

    The function returns two lists: window handles for all opened course
    windows and the corresponding course IDs in the order discovered.  If
    no unpassed courses are found, both lists will be empty.
    """
    wait = WebDriverWait(driver, 20)

    # Attempt to reveal a hidden user menu by clicking on the displayed user
    # name (which often appears as a surname followed by asterisks).  This
    # step helps ensure that the "我修的課" link becomes visible.
    try:
        user_elem = driver.find_element(By.XPATH, "//*[contains(text(), '＊')]")
        user_elem.click()
    except Exception:
        try:
            user_elem = driver.find_element(By.XPATH, "//*[contains(text(), '*')]")
            user_elem.click()
        except Exception:
            pass

    # Click on the "我修的課" link.  If it isn't immediately clickable,
    # try opening the user dropdown or hamburger menu to reveal it.
    try:
        my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
    except Exception:
        try:
            user_menu = driver.find_element(
                By.XPATH,
                "//*[(contains(text(), '＊') or contains(text(), '*')) and (self::span or self::button or self::a)]",
            )
            user_menu.click()
            my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
        except Exception:
            try:
                hamburger = driver.find_element(
                    By.CSS_SELECTOR,
                    "button.navbar-toggler, button.hamburger, button[aria-label='Toggle navigation']",
                )
                hamburger.click()
                my_courses_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "我修的課")))
            except Exception as e:
                print(f"[Navigate] Unable to find '我修的課': {e}")
                return [], []

    my_courses_link.click()
    # Wait briefly for the course list to render
    time.sleep(3)

    course_windows: List[str] = []
    course_ids: List[str] = []
    seen_ids = set()
    parent_handle = driver.current_window_handle

    # Collect and log course titles currently visible on the page for debugging.
    # We gather titles from the header rows (table_accordion-head) rather than from
    # status labels, since the status text may not be directly accessible via text().
    course_titles: List[str] = []
    try:
        header_rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'table_accordion-head')]")
    except Exception:
        header_rows = []
    for row in header_rows:
        try:
            # Course names are displayed in a <p> element with class containing 'course-name'.
            title_elem = row.find_element(By.XPATH, ".//p[contains(@class, 'course-name')]")
            title_text = title_elem.text.strip()
            if title_text:
                course_titles.append(title_text)
        except Exception:
            # If the expected <p> element isn't found, attempt to retrieve from an <a> tag.
            try:
                title_elem_alt = row.find_element(By.XPATH, ".//a")
                title_text = title_elem_alt.text.strip()
                if title_text:
                    course_titles.append(title_text)
            except Exception:
                continue
    if course_titles:
        print(f"[Navigate] 課程列表：{course_titles}")

    # Process each course with "進行中" status sequentially.  We skip any
    # duplicates by checking the course ID against the seen_ids set.  We do
    # not prematurely break out of the loop when encountering duplicates; instead
    # we continue scanning until a new course is found or no candidates remain.
    while True:
        processed = False
        try:
            status_elems = driver.find_elements(By.XPATH, "//*[contains(text(), '進行中')]")
        except Exception:
            status_elems = []
        if not status_elems:
            break
        for status_elem in status_elems:
            # Locate the course link associated with this status element.
            try:
                course_link = status_elem.find_element(By.XPATH, "ancestor::tr[1]//a")
            except Exception:
                try:
                    course_link = status_elem.find_element(By.XPATH, "..//a")
                except Exception:
                    continue
            # Extract the course title for logging (optional)
            course_title = course_link.text.strip()
            # Click the course link to navigate to the course page.
            try:
                course_link.click()
            except Exception:
                continue
            # Wait for the course page to load (URL includes '/learning/').
            try:
                WebDriverWait(driver, 15).until(lambda d: "/learning/" in d.current_url)
            except Exception:
                # If navigation fails, attempt to go back and continue.
                try:
                    driver.back()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
                except Exception:
                    pass
                continue
            course_url = driver.current_url
            # Extract numeric course ID from the URL.
            match = re.search(r"(\d+)$", course_url)
            course_id = match.group(1) if match else course_url
            if course_id in seen_ids:
                # Already processed; navigate back and skip this element.
                try:
                    driver.back()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
                except Exception:
                    pass
                continue
            seen_ids.add(course_id)
            course_ids.append(course_id)
            # Open the course URL in a new window.
            try:
                driver.switch_to.new_window('window')
                driver.get(course_url)
                new_handle = driver.current_window_handle
            except Exception:
                driver.execute_script("window.open(arguments[0], '_blank');", course_url)
                new_handle = driver.window_handles[-1]
                driver.switch_to.window(new_handle)
            course_windows.append(new_handle)
            print(f"[Navigate] Opened course window {new_handle} for course ID {course_id} (課程名稱：{course_title}).")
            # Return to the parent window and go back to the list.
            try:
                driver.switch_to.window(parent_handle)
                driver.back()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), '進行中')]")))
            except Exception:
                driver.get('https://moocs.moe.edu.tw/moocs/#/my-learning')
                time.sleep(3)
            processed = True
            # After processing one new course, break to re-fetch status elements
            break
        if not processed:
            # No new course processed in this iteration; exit the loop.
            break

    if not course_windows:
        print("[Navigate] 沒有找到標示為未完成的課程。")

    return course_windows, course_ids


def main() -> None:
    """Entry point for the automation script.

    This function parses command‑line arguments, configures the Chrome
    WebDriver (optionally in headless mode), prompts the user to log in,
    collects in‑progress courses, opens each in a new window, and starts
    background threads to handle periodic tab toggling.
    """
    parser = argparse.ArgumentParser(description="Automate interactions with the MOE MOOCs platform")
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run Chrome in headless mode (no GUI). Useful for server environments."
    )
    # When this flag is set, each course window will be opened in its own
    # headless Chrome instance while the main driver remains in its
    # original mode (headless or not).  This is useful if you want to
    # watch the main page but keep background course windows hidden.
    parser.add_argument(
        "--child-headless",
        action="store_true",
        help="Open child course windows in headless mode while keeping the main driver visible."
    )
    args = parser.parse_args()

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    if args.headless:
        options.add_argument('--headless')
        # Headless Chrome often requires a defined window size
        options.add_argument('--window-size=1920,1080')

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get('https://moocs.moe.edu.tw/moocs/#/home')

    print("請在瀏覽器中登入磨課師平台，完成後回到終端機並按 Enter。")
    input()

    # Open in‑progress courses and capture their drivers/handles.  The
    # child_headless flag determines whether each course is opened in
    # its own headless driver or in a new window of the main driver.
    course_pairs, course_ids = open_in_progress_courses_mod(driver, child_headless=args.child_headless)
    if not course_pairs:
        print("沒有找到未完成課程，或是打開課程時發生錯誤。請確認您已登入且有未完成課程。")
        return

    # Launch a thread for each course window.  Maintain a set of all
    # drivers (main and child) so we can close them on exit.  When
    # using child headless mode, each pair contains a distinct driver.
    threads: List[threading.Thread] = []
    drivers_to_close = set()
    drivers_to_close.add(driver)
    for i, (course_driver, handle) in enumerate(course_pairs):
        drivers_to_close.add(course_driver)
        thread = threading.Thread(target=click_periodically, args=(course_driver, handle), daemon=True)
        threads.append(thread)
        thread.start()
        print(f"[Main] Started click loop thread for window {handle}.")
        # Add startup interval for child-headless mode to avoid resource competition
        if args.child_headless and i < len(course_pairs) - 1:
            print(f"[Main] Waiting 3 seconds before starting next thread...")
            time.sleep(3)

    print(f"已開始對以下課程進行自動操作（順序為找到的順序）：{course_ids}")
    print("按 Ctrl+C 可中止。")
    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        print("接收到中斷訊號，關閉瀏覽器並結束程式…")
    finally:
        # Close all drivers (main and child) gracefully
        for drv in drivers_to_close:
            try:
                drv.quit()
            except Exception:
                pass


if __name__ == '__main__':
    main()