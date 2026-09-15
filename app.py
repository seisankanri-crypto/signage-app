import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import requests
import time
import os

# ==========================================
# 設定
# ==========================================
SPREADSHEET_ID = "1ThtSEj2dnEYSKHIXerjcujo9-6rxmRXEYk2ijS80OlQ"
COMPANY_NAME = "株式会社ハイビックス"
SLIDE_INTERVAL = 10  # スライド切替秒数
WEATHER_CITY = "Mizuho, Gifu, JP"  # 天気取得用

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(
    page_title=f"{COMPANY_NAME} インフォメーション",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# スタイル設定
# ==========================================
st.markdown("""
    <style>
        .stApp {
            background-color: #0d1117;
        }
        section.main > div {
            background-color: #0d1117;
            padding: 10px 20px;
        }
        .company-header {
            background: linear-gradient(135deg, #1a237e, #1565c0, #0288d1);
            padding: 10px;
            border-radius: 15px;
            text-align: center;
            margin-bottom: 10px;
            box-shadow: 0 4px 20px rgba(0, 100, 255, 0.4);
        }
        .company-name {
            color: white;
            font-size: 28px;
            font-weight: bold;
            letter-spacing: 2px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        .datetime-display {
            color: #90caf9;
            font-size: 18px;
            margin-top: 4px;
        }
        .slide-title {
            background: linear-gradient(135deg, #1565c0, #0288d1);
            color: white;
            padding: 10px 20px;
            border-radius: 12px;
            font-size: 22px;
            font-weight: bold;
            margin-bottom: 10px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0, 100, 255, 0.3);
            letter-spacing: 1px;
        }
        .notice-box {
            background: linear-gradient(135deg, #1e2a3a, #1a237e);
            border-left: 5px solid #0288d1;
            border-radius: 12px;
            padding: 12px 20px;
            margin: 8px 0;
            color: #ffffff;
            font-size: 16px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
        }
        .weather-box {
            background: linear-gradient(135deg, #1565c0, #0288d1, #00bcd4);
            border-radius: 20px;
            padding: 20px;
            text-align: center;
            color: white;
            box-shadow: 0 8px 30px rgba(0, 150, 255, 0.4);
        }
        .stDataFrame {
            font-size: 14px !important;
            border-radius: 12px !important;
        }
        .empty-message {
            text-align: center;
            color: #546e7a;
            font-size: 18px;
            padding: 40px 0;
        }
        .footer {
            text-align: center;
            color: #37474f;
            font-size: 14px;
            margin-top: 10px;
            letter-spacing: 3px;
        }
        ::-webkit-scrollbar {
            display: none;
        }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .main .block-container {
            overflow: hidden;
            max-height: 100vh;
            padding-top: 10px;
            padding-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 日付正規化関数 ← 追加
# ==========================================
def normalize_date(date_str):
    """日付を正規化する（2024/1/15 → 2024/01/15）"""
    try:
        parts = str(date_str).split("/")
        if len(parts) == 3:
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
            return f"{year:04d}/{month:02d}/{day:02d}"
    except:
        pass
    return date_str

# ==========================================
# スプレッドシート接続
# ==========================================
@st.cache_resource
def get_gspread_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )

    return gspread.Client(auth=creds)


def get_sheet_data(sheet_name):
    try:
        client = get_gspread_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet = spreadsheet.worksheet(sheet_name)
        data = sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"スプレッドシート読み込みエラー: {e}")
        st.error(f"詳細: {type(e).__name__}")
        return pd.DataFrame()


# ==========================================
# 天気情報取得
# ==========================================
def get_weather():
    try:
        lat = 35.3833
        lon = 136.7167
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&current_weather=true&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=Asia%2FTokyo&forecast_days=1"
        response = requests.get(url, timeout=5)
        data = response.json()

        current = data["current_weather"]
        daily = data["daily"]

        temp = current["temperature"]
        temp_max = daily["temperature_2m_max"][0]
        temp_min = daily["temperature_2m_min"][0]
        weathercode = current["weathercode"]

        weather_map = {
            0: ("☀️", "快晴"),
            1: ("🌤️", "晴れ"),
            2: ("⛅", "曇りがち"),
            3: ("☁️", "曇り"),
            45: ("🌫️", "霧"),
            48: ("🌫️", "霧"),
            51: ("🌦️", "小雨"),
            53: ("🌧️", "雨"),
            55: ("🌧️", "強い雨"),
            61: ("🌧️", "雨"),
            63: ("🌧️", "雨"),
            65: ("🌧️", "大雨"),
            71: ("❄️", "雪"),
            73: ("❄️", "雪"),
            75: ("❄️", "大雪"),
            80: ("🌦️", "にわか雨"),
            81: ("🌧️", "にわか雨"),
            82: ("⛈️", "激しい雨"),
            95: ("⛈️", "雷雨"),
            99: ("⛈️", "激しい雷雨"),
        }
        icon, desc = weather_map.get(weathercode, ("🌡️", "不明"))

        return {
            "icon": icon,
            "description": desc,
            "temp": temp,
            "temp_max": temp_max,
            "temp_min": temp_min
        }
    except Exception:
        return None

# ==========================================
# ヘッダー表示
# ==========================================
def show_header():
    now = datetime.now()
    date_str = now.strftime("%Y年%m月%d日（" + "月火水木金土日"[now.weekday()] + "）")
    time_str = now.strftime("%H:%M")

    st.markdown(f"""
        <div class="company-header">
            <div class="company-name"> {COMPANY_NAME}</div>
            <div class="datetime-display">{date_str}　{time_str}</div>
        </div>
    """, unsafe_allow_html=True)

# ==========================================
# スライド1：欠勤者一覧
# ==========================================
def show_absence():
    st.markdown('<div class="slide-title">🏥 本日の欠勤・遅刻・早退者</div>', unsafe_allow_html=True)

    df = get_sheet_data("欠勤連絡")

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    # 今日の日付でフィルタ ← 正規化して比較
    today = date.today().strftime("%Y/%m/%d")
    if "日付" in df.columns:
        df["日付_正規化"] = df["日付"].apply(normalize_date)
        df = df[df["日付_正規化"] == today]

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    display_cols = ["氏名", "部署", "種別", "備考", "登録時刻"]
    df_display = df[[col for col in display_cols if col in df.columns]]

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        height=400
    )

# ==========================================
# スライド2：来客情報
# ==========================================
def show_visitors():
    st.markdown('<div class="slide-title">🤝 本日の来客情報</div>', unsafe_allow_html=True)

    df = get_sheet_data("来客情報")

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    # 今日の日付でフィルタ ← 正規化して比較
    today = date.today().strftime("%Y/%m/%d")
    if "日付" in df.columns:
        df["日付_正規化"] = df["日付"].apply(normalize_date)
        df = df[df["日付_正規化"] == today]

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    display_cols = ["来訪時刻", "会社名", "訪問者名", "人数", "担当者"]
    df_display = df[[col for col in display_cols if col in df.columns]]

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        height=400
    )

# ==========================================
# スライド3：お知らせ
# ==========================================
def show_notices():
    st.markdown('<div class="slide-title">📢 お知らせ</div>', unsafe_allow_html=True)

    df = get_sheet_data("お知らせ")

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    # 掲載期間でフィルタ ← 正規化して比較
    today = date.today().strftime("%Y/%m/%d")
    if "掲載開始日" in df.columns and "掲載終了日" in df.columns:
        df["掲載開始日_正規化"] = df["掲載開始日"].apply(normalize_date)
        df["掲載終了日_正規化"] = df["掲載終了日"].apply(normalize_date)
        df = df[
            (df["掲載開始日_正規化"] <= today) &
            (df["掲載終了日_正規化"] >= today)
        ]

    if df.empty:
        st.markdown("""
            <div class="empty-message">
                📭 本日の連絡はありません
            </div>
        """, unsafe_allow_html=True)
        return

    for _, row in df.iterrows():
        st.markdown(f"""
            <div class="notice-box">
                📌 {row.get("お知らせ内容", "")}
                <br><small style="color:#666;">掲載期間：{row.get("掲載開始日","")} ～ {row.get("掲載終了日","")}</small>
            </div>
        """, unsafe_allow_html=True)

# ==========================================
# スライド4：天気情報（今週の天気）
# ==========================================
def show_weather():
    st.markdown('<div class="slide-title">🌤️ 今週の天気　岐阜県瑞穂市</div>', unsafe_allow_html=True)

    try:
        lat = 35.3833
        lon = 136.7167
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&daily=weathercode,temperature_2m_max,temperature_2m_min&timezone=Asia%2FTokyo&forecast_days=7"
        response = requests.get(url, timeout=5)
        data = response.json()

        daily = data["daily"]
        dates = daily["time"]
        weathercodes = daily["weathercode"]
        temp_maxs = daily["temperature_2m_max"]
        temp_mins = daily["temperature_2m_min"]

        weather_map = {
            0: ("☀️", "快晴"),
            1: ("🌤️", "晴れ"),
            2: ("⛅", "曇りがち"),
            3: ("☁️", "曇り"),
            45: ("🌫️", "霧"),
            48: ("🌫️", "霧"),
            51: ("🌦️", "小雨"),
            53: ("🌧️", "雨"),
            55: ("🌧️", "強い雨"),
            61: ("🌧️", "雨"),
            63: ("🌧️", "雨"),
            65: ("🌧️", "大雨"),
            71: ("❄️", "雪"),
            73: ("❄️", "雪"),
            75: ("❄️", "大雪"),
            80: ("🌦️", "にわか雨"),
            81: ("🌧️", "にわか雨"),
            82: ("⛈️", "激しい雨"),
            95: ("⛈️", "雷雨"),
            99: ("⛈️", "激しい雷雨"),
        }

        # 曜日リスト
        youbi = ["月", "火", "水", "木", "金", "土", "日"]

        # 7日分のカードを横並びで表示
        cols = st.columns(7)
        for i, col in enumerate(cols):
            d = datetime.strptime(dates[i], "%Y-%m-%d")
            youbi_str = youbi[d.weekday()]
            date_str = d.strftime(f"%m/%d\n（{youbi_str}）")
            icon, desc = weather_map.get(weathercodes[i], ("🌡️", "不明"))
            temp_max = temp_maxs[i]
            temp_min = temp_mins[i]

            # 今日は強調表示
            is_today = dates[i] == date.today().strftime("%Y-%m-%d")
            bg_color = "rgba(2, 136, 209, 0.5)" if is_today else "rgba(21, 101, 192, 0.2)"
            border = "2px solid #0288d1" if is_today else "1px solid #1565c0"

            col.markdown(f"""
                <div style="
                    background: {bg_color};
                    border: {border};
                    border-radius: 15px;
                    padding: 12px 5px;
                    text-align: center;
                    color: white;
                    box-shadow: 0 4px 15px rgba(0, 100, 255, 0.3);
                ">
                    <div style="font-size:13px; color:#90caf9; font-weight:bold;">
                        {d.strftime("%m/%d")}<br>（{youbi_str}）
                    </div>
                    <div style="font-size:36px; margin:6px 0;">{icon}</div>
                    <div style="font-size:12px; color:#e0e0e0;">{desc}</div>
                    <div style="font-size:16px; color:#ff7043; font-weight:bold; margin-top:6px;">
                        ▲ {temp_max}°C
                    </div>
                    <div style="font-size:16px; color:#42a5f5; font-weight:bold;">
                        ▼ {temp_min}°C
                    </div>
                </div>
            """, unsafe_allow_html=True)

    except Exception as e:
        st.markdown("""
            <div class="empty-message">
                📭 天気情報を取得できませんでした
            </div>
        """, unsafe_allow_html=True)

# ==========================================
# メイン処理
# ==========================================
def main():
    if "slide_index" not in st.session_state:
        st.session_state.slide_index = 0
    if "last_switch" not in st.session_state:
        st.session_state.last_switch = time.time()

    now = time.time()
    if now - st.session_state.last_switch >= SLIDE_INTERVAL:
        st.session_state.slide_index = (st.session_state.slide_index + 1) % 4
        st.session_state.last_switch = now

    show_header()

    slide = st.session_state.slide_index
    if slide == 0:
        show_absence()
    elif slide == 1:
        show_visitors()
    elif slide == 2:
        show_notices()
    elif slide == 3:
        show_weather()

    indicators = ""
    for i in range(4):
        if i == slide:
            indicators += "⬤ "
        else:
            indicators += "○ "

    st.markdown(f"""
        <div class="footer">
            {indicators}<br>
            最終更新：{datetime.now().strftime("%H:%M:%S")}
        </div>
    """, unsafe_allow_html=True)

    time.sleep(3)
    st.rerun()

if __name__ == "__main__":
    main()
