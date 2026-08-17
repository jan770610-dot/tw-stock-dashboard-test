# 台股分析中心・Streamlit Cloud 測試版

這是用來驗證「手機不在同一個 Wi‑Fi，也可以透過 Internet 開啟台股分析網頁」的最小測試版。

## GitHub 需要上傳
- `app.py`
- `requirements.txt`

`.gitignore` 與本 README 可一起上傳，但不是必要條件。

## 部署
1. 在 GitHub 建立 repository，例如 `tw-stock-dashboard-test`。
2. 將本資料夾內檔案上傳到 repository 根目錄。
3. 登入 Streamlit Community Cloud。
4. 建立新 App，選擇剛才的 repository。
5. Main file path 填 `app.py`。
6. 按 Deploy。
7. 等部署完成後，會取得 `https://xxxxx.streamlit.app` 類型網址。
8. 手機關閉 Wi‑Fi，改用 4G/5G，再開該網址。
9. 按頁面上的「按這裡測試遠端操作」。

## 測試成功判定
- 手機 4G/5G 可以開頁面。
- 可以看到「台股分析中心」。
- 按測試按鈕後，成功次數會增加。

這三項都成功，就代表網頁遠端使用路徑已成立。
