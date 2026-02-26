# learningTool
---
## 環境安裝

```bash
# 建立虛擬環境
python3 -m venv venv

# 啟動虛擬環境（bash/zsh）
source venv/bin/activate

# 安裝 selenium
python3 -m pip install selenium
```

---

## 使用方式

```bash
# 一般模式（登入與自動化都在有 GUI 的視窗中進行）
python3 mooc_auto.py

# --headless 模式（登入用 GUI 視窗，登入確認後切換至 headless 背景執行）
python3 mooc_auto.py --headless

# --child-headless 模式（主視窗維持 GUI，每門課的子視窗在背景 headless 中執行）
python3 mooc_auto.py --child-headless
```

---

## 執行流程

1. 瀏覽器自動開啟並切換語言為繁體中文
2. 使用者在瀏覽器中完成登入（包含 CAPTCHA）
3. 回到終端機按 **Enter**，腳本確認登入狀態
4. 自動導向「我修的課」，列出所有進行中的課程並各自開啟視窗
5. 每 30 秒循序檢查各課程視窗：切換「通過標準」與「課程簡介」分頁，並讀取閱讀時數進度
6. 某門課達到 100% 且分鐘數可被 5 整除時，自動關閉該視窗並移除追蹤
7. 所有課程完成後程式自動結束；按 **Ctrl+C** 可隨時中止

---

## 注意事項

- 需要安裝與 Chrome 版本對應的 [ChromeDriver](https://chromedriver.chromium.org/)，或使用 `selenium-manager`（selenium 4.6+ 自動管理）
- `--headless` 模式僅轉移 HTTP cookies；若平台將登入狀態存於 localStorage，headless driver 可能仍顯示為未登入，此時請改用一般模式
