from datetime import datetime
import streamlit as st

st.set_page_config(
    page_title="台股分析中心・Cloud 測試版",
    page_icon="📈",
    layout="wide",
)

# Streamlit Community Cloud 適用版：
# 不依賴本機檔案寫入。按鈕測試狀態保留在目前瀏覽工作階段內。
if "count" not in st.session_state:
    st.session_state.count = 0
if "last_click" not in st.session_state:
    st.session_state.last_click = "尚未測試"


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


st.title("📈 台股分析中心")
st.caption("Streamlit Community Cloud 測試版｜目前資料皆為示範，不是實際盤後分析結果")
st.success("如果你現在是用手機 4G/5G 開啟這個網址，代表『不同 Wi‑Fi 也能使用』的第一階段已成功。")

c1, c2, c3, c4 = st.columns(4)
c1.metric("大趨勢", "測試模式")
c2.metric("市場狀態", "雲端連線正常")
c3.metric("鴨嘴候選", "3 檔", help="純示範數字")
c4.metric("RS 強勢股", "12 檔", help="純示範數字")

st.divider()

left, right = st.columns(2)
with left:
    st.subheader("🦆 鴨嘴系統")
    st.write("正式版可接入：均線、鴨嘴型態、日／週／月週期、主升段判定與個股篩選。")
    st.dataframe(
        {
            "代號": ["TEST1", "TEST2", "TEST3"],
            "狀態": ["示範候選", "示範候選", "示範候選"],
            "週期": ["日線", "週線", "月線"],
        },
        use_container_width=True,
        hide_index=True,
    )

with right:
    st.subheader("📊 RS／市場廣度")
    st.write("正式版可接入：RS 強勢股數量、市場廣度、極值趨勢、大趨勢與市場狀態。")
    st.dataframe(
        {
            "項目": ["RS 強勢股", "創新高家數", "市場廣度"],
            "示範值": [12, 8, 68],
            "說明": ["測試", "測試", "測試"],
        },
        use_container_width=True,
        hide_index=True,
    )

st.divider()
st.subheader("📱 手機遠端操作測試")
st.write(
    f"本次工作階段上次按鈕時間：**{st.session_state.last_click}**　｜　"
    f"累計成功：**{st.session_state.count} 次**"
)

if st.button("✅ 按這裡測試遠端操作", type="primary", use_container_width=True):
    st.session_state.count += 1
    st.session_state.last_click = now_text()
    st.success(f"成功！雲端 App 已收到你的操作：{st.session_state.last_click}")

st.divider()
st.info(
    "下一階段：把你現有的『鴨嘴系統』與『RS 強勢股／市場廣度』實際計算結果接進這個頁面，"
    "再加入『立即更新』、Excel 下載與手機版檢視。"
)
st.caption("注意：免費雲端測試版的工作階段可能重新啟動，因此這裡的按鈕次數只用來驗證互動是否成功，不作永久保存。")
