import streamlit as st
import datetime
import time
import base64

# ページ設定
st.set_page_config(
    page_title="社内サイネージ",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 画像をbase64に変換する関数
def get_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# 画像読み込み
bg_image = get_base64("syaoku.jpg")
logo_image = get_base64("IMG_2171.png")

# CSS設定
st.markdown(f"""
    <style>
    /* 全体背景 */
    .stApp {{
        background-image: url("data:image/jpg;base64,{bg_image}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
    }}
    
    /* 背景ぼかしオーバーレイ */
    .stApp::before {{
        content: "";
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: url("data:image/jpg;base64,{bg_image}");
        background-size: cover;
        background-position: center;
        filter: blur(8px);
        z-index: -1;
    }}

    /* 暗めのオーバーレイ（文字を見やすく） */
    .overlay {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.4);
        z-index: -1;
    }}

    /* ロゴ */
    .logo {{
        position: fixed;
        bottom: 20px;
        right: 20px;
        width: 150px;
        z-index: 100;
    }}

    /* 日時 */
    .datetime {{
        color: white;
        font-size: 60px;
        font-weight: bold;
        text-align: center;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.8);
        margin-top: 50px;
    }}

    /* お知らせタイトル */
    .notice-title {{
        color: white;
        font-size: 36px;
        font-weight: bold;
        text-align: center;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.8);
        margin-top: 30px;
    }}

    /* お知らせ内容 */
    .notice-content {{
        color: white;
        font-size: 28px;
        text-align: center;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.8);
        margin-top: 10px;
        line-height: 1.8;
    }}

    /* Streamlitのデフォルト要素を非表示 */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}
    </style>

    <div class="overlay"></div>
    <img class="logo" src="data:image/png;base64,{logo_image}">
""", unsafe_allow_html=True)

# スプレッドシートからお知らせ取得
import gspread
from google.oauth2.service_account import Credentials
import json

def get_notices():
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(st.secrets["SPREADSHEET_ID"]).worksheet("お知らせ")
        records = sheet.get_all_records()
        today = datetime.date.today()
        notices = []
        for record in records:
            try:
                start = datetime.datetime.strptime(str(record["掲載開始日"]), "%Y-%m-%d").date()
                end = datetime.datetime.strptime(str(record["掲載終了日"]), "%Y-%m-%d").date()
                if start <= today <= end:
                    notices.append(record["お知らせ内容"])
            except:
                continue
        return notices
    except:
        return []

# メインループ
placeholder = st.empty()

while True:
    now = datetime.datetime.now()
    date_str = now.strftime("%Y年%m月%d日")
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    weekday_str = weekdays[now.weekday()]
    time_str = now.strftime("%H:%M")
    
    notices = get_notices()
    
    with placeholder.container():
        # 日時表示
        st.markdown(f"""
            <div class="datetime">
                {date_str}({weekday_str})　{time_str}
            </div>
        """, unsafe_allow_html=True)
        
        # お知らせ表示
        if notices:
            st.markdown('<div class="notice-title">📢 お知らせ</div>', 
                       unsafe_allow_html=True)
            for notice in notices:
                st.markdown(f"""
                    <div class="notice-content">
                        {notice}
                    </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div class="notice-title">本日のお知らせはありません</div>
            """, unsafe_allow_html=True)
    
    time.sleep(60)
