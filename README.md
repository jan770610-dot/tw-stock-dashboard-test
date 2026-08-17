# 台股分析中心 v2

這一版加入：
- 手機友善的 KPI 卡片，不再把「偏低水位」「廣度差」截成省略號。
- RS 圖表明確圖例。
- 自動更新狀態。
- GitHub Actions 週一至週五台灣時間 18:05 自動更新。
- 官方資料尚未齊時，8 分鐘後重試，最多 3 次。
- 首次執行以 2026-08-14 的既有 stock.db / RS cache 當種子，之後用 Actions cache 做增量更新。
- Actions 頁面可手動執行，也可指定日期。

重要：目前 repository 是 Public，所以上傳 v2 後，automation 內的策略程式與 seed 資料也會公開可見。若要保護策略程式，後續應改成 private automation repo 或完成 Streamlit private repo 授權。
