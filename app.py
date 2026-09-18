import streamlit as st
from streamlit_autorefresh import st_autorefresh
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import requests
import time
import os
import pytz
import base64

# ==========================================
# 設定
# ==========================================
SPREADSHEET_ID = "1ThtSEj2dnEYSKHIXerjcujo9-6rxmRXEYk2ijS80OlQ"
COMPANY_NAME = "株式会社ハイビックス"
SLIDE_INTERVAL = 10
WEATHER_CITY = "Mizuho, Gifu, JP"
JST = pytz.timezone("Asia/Tokyo")

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(
    page_title=f"{COMPANY_NAME} インフォメーション",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 自動更新（チラつき防止）
# ==========================================
st_autorefresh(interval=3000, key="autorefresh")

# ==========================================
# 画像をbase64に変換 ✅ キャッシュして1回だけ読み込む
# ==========================================
@st.cache_resource
def get_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

bg_image = get_base64("syaoku2.jpg")
logo_image = get_base64("IMG_2185 (2).png")

# ==========================================
# スタイル設定
# ==========================================
st.markdown(f"""
    <style>
        .stApp {{
            background-image: url("data:image/jpg;base64,{bg_image}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
        }}
        .stApp::before {{
            content: "";
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.45);
            z-index: 0;
        }}
        section.main > div {{
            position: relative;
            z-index: 1;
            padding: 10px 20px;
        }}
        .logo-fixed {{
            position: fixed;
            bottom: 30px;
            left: 20px;
            width: 280px;
            z-index: 9999;
        }}
        .company-header {{
            background: linear-gradient(135deg, rgba(26,35,126,0.85), rgba(21,101,192,0.85), rgba(2,136,209,0.85));
            padding: 10px;
            border-radius: 15px;
            text-align: center;
            margin-bottom: 10px;
            box-shadow: 0 4px 20px rgba(0, 100, 255, 0.4);
        }}
        .company-name {{
            color: white;
            font-size: 28px;
            font-weight: bold;
            letter-spacing: 2px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }}
        .datetime-display {{
            color: #90caf9;
            font-size: 18px;
            margin-top: 4px;
        }}
        .slide-title {{
            background: linear-gradient(135deg, rgba(21,101,192,0.85), rgba(2,136,209,0.85));
            color: white;
            padding: 10px 20px;
            border-radius: 12px;
            font-size: 22px;
            font-weight: bold;
            margin-bottom: 10px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0, 100, 255, 0.3);
            letter-spacing: 1px;
        }}
        .notice-box {{
            background: linear-gradient(135deg, rgba(30,42,58,0.85), rgba(26,35,126,0.85));
            border-left: 5px solid #0288d1;
            border-radius: 12px;
            padding: 12px 20px;
            margin: 8px 0;
            color: #ffffff;
            font-size: 16px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
        }}
        .empty-message {{
            text-align: center;
            color: #ffffff;
            font-size: 18px;
            padding: 40px 0;
            text-shadow: 1px 1px 3px rgba(0,0,0,0.8);
        }}
        .footer {{
            text-align: center;
            color: #ffffff;
            font-size: 14px;
            margin-top: 10px;
            letter-spacing: 3px;
            text-shadow: 1px 1px 3px rgba(0,0,0,0.8);
        }}
        ::-webkit-scrollbar {{
            display: none;
        }}
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}
        .main .block-container {{
            overflow: hidden;
            max-height: 100vh;
            padding-top: 10px;
            padding-bottom: 10px;
        }}
        .stDataFrame {{
            font-size: 14px !important;
            border-radius: 12px !important;
        }}
        .custom-table {{
            width: 100%;
            border-collapse: collapse;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0,0,0,0.4);
            font-size: 16px;
        }}
        .custom-table th {{
            background: linear-gradient(135deg, rgba(21,101,192,0.9), rgba(2,136,209,0.9));
            color: white;
            padding: 14px 16px;
            text-align: center;
            font-weight: bold;
            letter-spacing: 1px;
            border: none;
        }}
        .custom-table td {{
            padding: 12px 16px;
            text-align: center;
            color: white;
            border: none;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }}
        .row-even {{
            background: rgba(21, 101, 192, 0.3);
        }}
        .row-odd {{
            background: rgba(2, 136, 209, 0.15);
        }}
        .custom-table tr:hover td {{
            background: rgba(2, 136, 209, 0.5);
        }}
    </style>

    <!-- ロゴ固定表示 -->
    <img class="logo-fixed" src="data:image/png;base64,{logo_image}">
""", unsafe_allow_html=True)

# ==========================================
# 日付正規化関数
# ==========================================
def normalize_date(date_str):
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
        return pd.DataFrame()

# ==========================================
# ヘッダー表示
# ==========================================
def show_header():
    now = datetime.now(JST)
    date_str = now.strftime("%Y年%m月%d日（" + "月火水木金土日"[now.weekday()] + "）")
    time_str = now.strftime("%H:%M")
    st.markdown(f"""
        <div class="company-header">
            <div class="company-name">{COMPANY_NAME}</div>
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
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return
    today = datetime.now(JST).strftime("%Y/%m/%d")
    if "日付" in df.columns:
        df["日付_正規化"] = df["日付"].apply(normalize_date)
        df = df[df["日付_正規化"] == today]
    if df.empty:
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return

    # ✅ 有給休暇・振替休暇・欠勤は固定時間を表示
    TIME_FIXED_KINDS = ["有給休暇", "振替休暇", "欠勤"]
    def fill_time(row):
        if row.get("種別", "") in TIME_FIXED_KINDS:
            row["開始時刻"] = "8:30"
            row["終了時刻"] = "17:30"
        return row

    df = df.apply(fill_time, axis=1)

    # ✅ 備考・昼食は表示しない
    display_cols = ["氏名", "部署", "種別", "開始時刻", "終了時刻", "登録時刻"]
    df_display = df[[col for col in display_cols if col in df.columns]]

    headers = "".join([f"<th>{col}</th>" for col in df_display.columns])
    rows = ""
    for i, row in df_display.iterrows():
        cells = "".join([f"<td>{val}</td>" for val in row.values])
        row_class = "row-even" if i % 2 == 0 else "row-odd"
        rows += f'<tr class="{row_class}">{cells}</tr>'

    st.markdown(f"""
        <table class="custom-table">
            <thead><tr>{headers}</tr></thead>
            <tbody>{rows}</tbody>
        </table>
    """, unsafe_allow_html=True)

# ==========================================
# スライド2：来客情報
# ==========================================
def show_visitors():
    st.markdown('<div class="slide-title">🤝 本日の来客情報</div>', unsafe_allow_html=True)
    df = get_sheet_data("来客情報")
    if df.empty:
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return
    today = datetime.now(JST).strftime("%Y/%m/%d")
    if "日付" in df.columns:
        df["日付_正規化"] = df["日付"].apply(normalize_date)
        df = df[df["日付_正規化"] == today]
    if df.empty:
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return

    display_cols = ["来訪時刻", "会社名", "訪問者名", "人数", "担当者"]
    df_display = df[[col for col in display_cols if col in df.columns]]

    headers = "".join([f"<th>{col}</th>" for col in df_display.columns])
    rows = ""
    for i, row in df_display.iterrows():
        cells = "".join([f"<td>{val}</td>" for val in row.values])
        row_class = "row-even" if i % 2 == 0 else "row-odd"
        rows += f'<tr class="{row_class}">{cells}</tr>'

    st.markdown(f"""
        <table class="custom-table">
            <thead><tr>{headers}</tr></thead>
            <tbody>{rows}</tbody>
        </table>
    """, unsafe_allow_html=True)

# ==========================================
# スライド3：お知らせ
# ==========================================
def show_notices():
    st.markdown('<div class="slide-title">📢 お知らせ</div>', unsafe_allow_html=True)
    df = get_sheet_data("お知らせ")
    if df.empty:
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return
    today = datetime.now(JST).strftime("%Y/%m/%d")
    if "掲載開始日" in df.columns and "掲載終了日" in df.columns:
        df["掲載開始日_正規化"] = df["掲載開始日"].apply(normalize_date)
        df["掲載終了日_正規化"] = df["掲載終了日"].apply(normalize_date)
        df = df[
            (df["掲載開始日_正規化"] <= today) &
            (df["掲載終了日_正規化"] >= today)
        ]
    if df.empty:
        st.markdown('<div class="empty-message">📭 本日の連絡はありません</div>', unsafe_allow_html=True)
        return
    for _, row in df.iterrows():
        st.markdown(f"""
            <div class="notice-box">
                📌 {row.get("お知らせ内容", "")}
                <br><small style="color:#aaa;">掲載期間：{row.get("掲載開始日","")} ～ {row.get("掲載終了日","")}</small>
            </div>
        """, unsafe_allow_html=True)

# ==========================================
# スライド4：天気情報
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
            0: ("☀️", "快晴"), 1: ("🌤️", "晴れ"), 2: ("⛅", "曇りがち"),
            3: ("☁️", "曇り"), 45: ("🌫️", "霧"), 48: ("🌫️", "霧"),
            51: ("🌦️", "小雨"), 53: ("🌧️", "雨"), 55: ("🌧️", "強い雨"),
            61: ("🌧️", "雨"), 63: ("🌧️", "雨"), 65: ("🌧️", "大雨"),
            71: ("❄️", "雪"), 73: ("❄️", "雪"), 75: ("❄️", "大雪"),
            80: ("🌦️", "にわか雨"), 81: ("🌧️", "にわか雨"),
            82: ("⛈️", "激しい雨"), 95: ("⛈️", "雷雨"), 99: ("⛈️", "激しい雷雨"),
        }
        youbi = ["月", "火", "水", "木", "金", "土", "日"]
        cols = st.columns(7)
        for i, col in enumerate(cols):
            d = datetime.strptime(dates[i], "%Y-%m-%d")
            youbi_str = youbi[d.weekday()]
            icon, desc = weather_map.get(weathercodes[i], ("🌡️", "不明"))
            today_str = datetime.now(JST).strftime("%Y-%m-%d")
            is_today = dates[i] == today_str
            bg_color = "rgba(2, 136, 209, 0.6)" if is_today else "rgba(21, 101, 192, 0.3)"
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
                    <div style="font-size:16px; color:#90caf9; font-weight:bold;">
                        {d.strftime("%m/%d")}<br>（{youbi_str}）
                    </div>
                    <div style="font-size:48px; margin:8px 0;">{icon}</div>
                    <div style="font-size:15px; color:#e0e0e0;">{desc}</div>
                    <div style="font-size:20px; color:#ff7043; font-weight:bold; margin-top:8px;">
                        ▲ {temp_maxs[i]}°C
                    </div>
                    <div style="font-size:20px; color:#42a5f5; font-weight:bold;">
                        ▼ {temp_mins[i]}°C
                    </div>
                </div>
            """, unsafe_allow_html=True)
    except Exception as e:
        st.markdown('<div class="empty-message">📭 天気情報を取得できませんでした</div>', unsafe_allow_html=True)

# ==========================================
# スライド5：ニュース
# ==========================================
def show_news():
    st.markdown('<div class="slide-title">📰 本日のニュース</div>', unsafe_allow_html=True)
    try:
        import xml.etree.ElementTree as ET
        url = "https://www.nhk.or.jp/rss/news/cat0.xml"
        response = requests.get(url, timeout=5)
        response.encoding = "utf-8"
        root = ET.fromstring(response.content)

        items = root.findall(".//item")[:8]

        news_html = ""
        for i, item in enumerate(items):
            title = item.find("title").text or ""
            pub_date = item.find("pubDate").text or ""

            try:
                dt = datetime.strptime(pub_date, "%a, %d %b %Y %H:%M:%S %z")
                dt_jst = dt.astimezone(JST)
                time_str = dt_jst.strftime("%H:%M")
            except:
                time_str = ""

            row_class = "rgba(21, 101, 192, 0.3)" if i % 2 == 0 else "rgba(2, 136, 209, 0.15)"

            news_html += f"""
                <div style="
                    background: {row_class};
                    border-left: 4px solid #0288d1;
                    border-radius: 8px;
                    padding: 12px 20px;
                    margin: 6px 0;
                    color: white;
                    display: flex;
                    align-items: center;
                    gap: 16px;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
                ">
                    <span style="
                        color: #90caf9;
                        font-size: 15px;
                        font-weight: bold;
                        min-width: 45px;
                    ">{time_str}</span>
                    <span style="font-size: 17px;">📌 {title}</span>
                </div>
            """

        st.markdown(news_html, unsafe_allow_html=True)

    except Exception as e:
        st.markdown('<div class="empty-message">📭 ニュースを取得できませんでした</div>', unsafe_allow_html=True)

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
        st.session_state.slide_index = (st.session_state.slide_index + 1) % 5
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
    elif slide == 4:
        show_news()

    indicators = ""
    for i in range(5):
        if i == slide:
            indicators += "⬤ "
        else:
            indicators += "○ "

    st.markdown(f"""
        <div class="footer">
            {indicators}<br>
            最終更新：{datetime.now(JST).strftime("%H:%M:%S")}
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
