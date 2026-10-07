import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date, timedelta
import pytz
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 設定
# ==========================================
SPREADSHEET_ID = "1ThtSEj2dnEYSKHIXerjcujo9-6rxmRXEYk2ijS80OlQ"
COMPANY_NAME = "株式会社ハイビックス"
ADMIN_PASSWORD = "3131"
JST = pytz.timezone("Asia/Tokyo")

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(
    page_title=f"{COMPANY_NAME} 管理画面",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# スタイル設定
# ==========================================
st.markdown("""
    <style>
        .stApp {
            background-color: #f0f4f8;
        }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
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

def get_sheet(sheet_name):
    client = get_gspread_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        # シートが無い場合は作成してヘッダーをセット
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=20)
        if sheet_name == "イベント一覧":
            ws.append_row(["イベントID", "イベント名", "受付状況", "質問内容", "選択肢1", "選択肢2", "選択肢3"])
        elif sheet_name == "イベント回答":
            ws.append_row(["イベントID", "イベント名", "氏名", "回答", "コメント", "回答日時"])
        return ws

def get_sheet_data(sheet_name):
    try:
        sheet = get_sheet(sheet_name)
        data = sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"読み込みエラー: {e}")
        return pd.DataFrame()

def delete_row(sheet_name, row_index):
    try:
        sheet = get_sheet(sheet_name)
        sheet.delete_rows(row_index + 2)
        return True
    except Exception as e:
        st.error(f"削除エラー: {e}")
        return False

# イベントID自動採番ロジック
def generate_event_id():
    df_events = get_sheet_data("イベント一覧")
    if df_events.empty or "イベントID" not in df_events.columns:
        return "EV01"
    
    ids = df_events["イベントID"].dropna().astype(str).tolist()
    num_list = []
    for eid in ids:
        if eid.startswith("EV"):
            try:
                num_list.append(int(eid.replace("EV", "")))
            except ValueError:
                pass
    if not num_list:
        return "EV01"
    
    max_num = max(num_list)
    return f"EV{max_num + 1:02d}"

# ==========================================
# ログイン画面
# ==========================================
def show_login():
    st.markdown("""
        <div style="text-align:center; margin-top:100px;">
            <h1>🏭 管理画面ログイン</h1>
            <p style="color:#666;">パスワードを入力してください</p>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password = st.text_input(
            "パスワード",
            type="password",
            placeholder="パスワードを入力",
        )
        login_btn = st.button("🔓 ログイン", use_container_width=True)

        if login_btn:
            if password == ADMIN_PASSWORD:
                st.session_state.logged_in = True
                st.success("ログイン成功！")
                st.rerun()
            else:
                st.error("パスワードが違います")

# ==========================================
# メイン処理
# ==========================================
def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        show_login()
        return

    st.title(f"🏭 {COMPANY_NAME} 管理画面")

    with st.sidebar:
        st.success("✅ ログイン中")
        if st.button("🔒 ログアウト", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

    menu = st.sidebar.selectbox(
        "メニューを選択",
        ["🏥 欠勤登録", "🤝 来客登録", "📢 お知らせ登録", "🎉 イベント管理", "📊 勤怠一覧", "🆘 安否確認一覧"]
    )

    # ==========================================
    # 欠勤登録
    # ==========================================
    if menu == "🏥 欠勤登録":
        st.header("🏥 欠勤・遅刻・早退 登録")

        absence_type = st.selectbox(
            "種別",
            ["欠勤", "有給休暇", "振替休暇", "休業", "遅刻", "早退", "私用外出"]
        )

        with st.form("absence_form"):
            col1, col2 = st.columns(2)

            with col1:
                name = st.text_input("氏名", placeholder="例：山田 太郎")
                department = st.text_input("部署", placeholder="例：製造部")

            with col2:
                if absence_type in ["欠勤", "有給休暇", "振替休暇", "休業"]:
                    start_date = st.date_input("開始日", value=datetime.now(JST).date())
                    end_date = st.date_input("終了日", value=datetime.now(JST).date())
                    start_time = None
                    end_time = None
                else:
                    start_date = st.date_input("日付", value=datetime.now(JST).date())
                    end_date = start_date
                    st.markdown("**開始時刻**")
                    start_time = st.time_input(
                        "開始時刻",
                        value=datetime.strptime("08:30", "%H:%M").time()
                    )
                    st.markdown("**終了時刻**")
                    end_time = st.time_input(
                        "終了時刻",
                        value=datetime.strptime("17:30", "%H:%M").time()
                    )

                note = st.text_area("備考", placeholder="例：発熱のため", height=80)
                lunch = st.selectbox("昼食", ["あり", "なし"])

            submitted = st.form_submit_button("✅ 登録する", use_container_width=True)

            if submitted:
                if not name or not department:
                    st.error("氏名と部署を入力してください")
                elif absence_type in ["欠勤", "有給休暇", "振替休暇", "休業"] and start_date > end_date:
                    st.error("終了日は開始日以降にしてください")
                else:
                    try:
                        sheet = get_sheet("欠勤連絡")
                        now_str = datetime.now(JST).strftime("%Y/%m/%d %H:%M")

                        date_list = []
                        current = start_date
                        while current <= end_date:
                            date_list.append(current)
                            current += timedelta(days=1)

                        if absence_type in ["欠勤", "有給休暇", "振替休暇", "休業"]:
                            start_time_str = "8:30"
                            end_time_str = "17:30"
                        else:
                            start_time_str = start_time.strftime("%H:%M")
                            end_time_str = end_time.strftime("%H:%M")

                        for d in date_list:
                            date_str = d.strftime("%Y/%m/%d")
                            sheet.append_row([
                                date_str,
                                name,
                                department,
                                absence_type,
                                start_time_str,
                                end_time_str,
                                note,
                                lunch,
                                now_str
                            ])

                        days = len(date_list)
                        st.success(f"✅ {name}さんの{absence_type}を{days}日分登録しました！")
                        st.cache_resource.clear()
                    except Exception as e:
                        st.error(f"登録エラー: {e}")

        st.subheader("📋 本日の登録済みデータ")
        df = get_sheet_data("欠勤連絡")

        if not df.empty and "日付" in df.columns:
            today = datetime.now(JST).strftime("%Y/%m/%d")
            df["日付_正規化"] = df["日付"].apply(normalize_date)
            df_today = df[df["日付_正規化"] == today].copy()
            df_today.reset_index(drop=False, inplace=True)

            if not df_today.empty:
                for _, row in df_today.iterrows():
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.info(
                            f"👤 {row.get('氏名','')} / "
                            f"{row.get('部署','')} / "
                            f"{row.get('種別','')} / "
                            f"{row.get('開始時刻','')}〜"
                            f"{row.get('終了時刻','')} / "
                            f"{row.get('備考','')}"
                        )
                    with col2:
                        if st.button("🗑️ 削除", key=f"absence_{row['index']}"):
                            if delete_row("欠勤連絡", row["index"]):
                                st.success("削除しました！")
                                st.cache_resource.clear()
                                st.rerun()
            else:
                st.info("本日の登録はありません")
        else:
            st.info("本日の登録はありません")

    # ==========================================
    # 来客登録
    # ==========================================
    elif menu == "🤝 来客登録":
        st.header("🤝 来客情報 登録")

        with st.form("visitor_form"):
            col1, col2 = st.columns(2)

            with col1:
                visit_date = st.date_input("日付", value=datetime.now(JST).date())
                visit_time = st.time_input("来訪時刻", value=datetime.now(JST).time())
                company = st.text_input("会社名", placeholder="例：〇〇株式会社")

            with col2:
                visitor_name = st.text_input("訪問者名", placeholder="例：鈴木 一郎")
                num_people = st.number_input("人数", min_value=1, max_value=50, value=1)
                person_in_charge = st.text_input("担当者", placeholder="例：田中 次郎")

            submitted = st.form_submit_button("✅ 登録する", use_container_width=True)

            if submitted:
                if not company or not visitor_name:
                    st.error("会社名と訪問者名を入力してください")
                else:
                    try:
                        sheet = get_sheet("来客情報")
                        date_str = visit_date.strftime("%Y/%m/%d")
                        time_str = visit_time.strftime("%H:%M")
                        sheet.append_row([
                            date_str,
                            time_str,
                            company,
                            visitor_name,
                            num_people,
                            person_in_charge
                        ])
                        st.success(f"✅ {company}様の来客情報を登録しました！")
                        st.cache_resource.clear()
                    except Exception as e:
                        st.error(f"登録エラー: {e}")

        st.subheader("📋 本日の登録済みデータ")
        df = get_sheet_data("来客情報")

        if not df.empty and "日付" in df.columns:
            today = datetime.now(JST).strftime("%Y/%m/%d")
            df["日付_正規化"] = df["日付"].apply(normalize_date)
            df_today = df[df["日付_正規化"] == today].copy()
            df_today.reset_index(drop=False, inplace=True)

            if not df_today.empty:
                for _, row in df_today.iterrows():
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.info(
                            f"🏢 {row.get('会社名','')} / "
                            f"{row.get('訪問者名','')} / "
                            f"{row.get('来訪時刻','')} / "
                            f"担当：{row.get('担当者','')}"
                        )
                    with col2:
                        if st.button("🗑️ 削除", key=f"visitor_{row['index']}"):
                            if delete_row("来客情報", row["index"]):
                                st.success("削除しました！")
                                st.cache_resource.clear()
                                st.rerun()
            else:
                st.info("本日の登録はありません")
        else:
            st.info("本日の登録はありません")

    # ==========================================
    # お知らせ登録
    # ==========================================
    elif menu == "📢 お知らせ登録":
        st.header("📢 お知らせ 登録")

        with st.form("notice_form"):
            content = st.text_area(
                "お知らせ内容",
                placeholder="例：〇〇のため、△△をお知らせします。",
                height=150
            )
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input("掲載開始日", value=datetime.now(JST).date())
            with col2:
                end_date = st.date_input("掲載終了日", value=datetime.now(JST).date())

            submitted = st.form_submit_button("✅ 登録する", use_container_width=True)

            if submitted:
                if not content:
                    st.error("お知らせ内容を入力してください")
                elif start_date > end_date:
                    st.error("掲載終了日は開始日以降にしてください")
                else:
                    try:
                        sheet = get_sheet("お知らせ")
                        start_str = start_date.strftime("%Y/%m/%d")
                        end_str = end_date.strftime("%Y/%m/%d")
                        sheet.append_row([content, start_str, end_str])
                        st.success("✅ お知らせを登録しました！")
                        st.cache_resource.clear()
                    except Exception as e:
                        st.error(f"登録エラー: {e}")

        st.subheader("📋 登録済みお知らせ一覧")
        df = get_sheet_data("お知らせ")

        if not df.empty:
            df_notice = df.copy()
            df_notice.reset_index(drop=False, inplace=True)
            for _, row in df_notice.iterrows():
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.info(
                        f"📌 {row.get('お知らせ内容','')} / "
                        f"{row.get('掲載開始日','')} ～ "
                        f"{row.get('掲載終了日','')}"
                    )
                with col2:
                    if st.button("🗑️ 削除", key=f"notice_{row['index']}"):
                        if delete_row("お知らせ", row["index"]):
                            st.success("削除しました！")
                            st.cache_resource.clear()
                            st.rerun()
        else:
            st.info("登録済みのお知らせはありません")

    # ==========================================
    # 🎉 イベント管理（新追加）
    # ==========================================
    elif menu == "🎉 イベント管理":
        st.header("🎉 イベント出欠・希望確認 管理")

        tab_create, tab_list, tab_answers = st.tabs([
            "➕ イベント新規作成",
            "📋 イベント一覧・受付切り替え",
            "📊 回答結果の確認"
        ])

        # --- タブ1: イベント新規作成 ---
        with tab_create:
            st.subheader("新しいイベントを作成する")
            auto_id = generate_event_id()
            st.info(f"🔑 発行されるイベントID: **{auto_id}** （自動採番）")

            with st.form("create_event_form"):
                event_name = st.text_input("イベント名", placeholder="例：忘年会出欠確認")
                question = st.text_area("質問内容（LINEで送信される本文）", placeholder="例：12/20(金)の忘年会に参加できますか？")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    opt1 = st.text_input("選択肢 1", placeholder="例：参加")
                with col2:
                    opt2 = st.text_input("選択肢 2", placeholder="例：不参加")
                with col3:
                    opt3 = st.text_input("選択肢 3（任意）", placeholder="例：検討中")

                status = st.selectbox("受付状況", ["受付中", "受付終了"])

                submitted = st.form_submit_button("🚀 イベントを作成する", use_container_width=True)

                if submitted:
                    if not event_name or not question or not opt1 or not opt2:
                        st.error("「イベント名」「質問内容」「選択肢1」「選択肢2」は必須です。")
                    else:
                        try:
                            sheet = get_sheet("イベント一覧")
                            sheet.append_row([
                                auto_id,
                                event_name,
                                status,
                                question,
                                opt1,
                                opt2,
                                opt3 if opt3 else ""
                            ])
                            st.success(f"✅ イベント「{event_name}」（ID: {auto_id}）を作成しました！")
                            st.cache_resource.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"作成エラー: {e}")

        # --- タブ2: イベント一覧・ステータス変更 ---
        with tab_list:
            st.subheader("登録済みイベント一覧")
            df_events = get_sheet_data("イベント一覧")

            if not df_events.empty and "イベントID" in df_events.columns:
                df_events_reset = df_events.copy()
                df_events_reset.reset_index(drop=False, inplace=True)

                for _, row in df_events_reset.iterrows():
                    status_badge = "🟢 受付中" if row.get("受付状況") == "受付中" else "🔴 受付終了"
                    
                    with st.expander(f"{status_badge} | 【{row.get('イベントID')}】{row.get('イベント名')}"):
                        st.write(f"**質問内容:** {row.get('質問内容')}")
                        st.write(f"**選択肢:** 1. {row.get('選択肢1')} / 2. {row.get('選択肢2')}" + (f" / 3. {row.get('選択肢3')}" if row.get('選択肢3') else ""))

                        col1, col2, col3 = st.columns([2, 2, 1])
                        
                        with col1:
                            # ステータス切り替えボタン
                            new_status = "受付終了" if row.get("受付状況") == "受付中" else "受付中"
                            if st.button(f"⚙️ ステータスを「{new_status}」に変更", key=f"status_btn_{row['index']}"):
                                try:
                                    sheet = get_sheet("イベント一覧")
                                    # 行番号は index + 2 (ヘッダー含む)
                                    sheet.update_cell(row['index'] + 2, 3, new_status)
                                    st.success(f"ステータスを「{new_status}」に変更しました！")
                                    st.cache_resource.clear()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"更新エラー: {e}")

                        with col3:
                            if st.button("🗑️ イベント削除", key=f"del_event_{row['index']}"):
                                if delete_row("イベント一覧", row["index"]):
                                    st.success("イベントを削除しました！")
                                    st.cache_resource.clear()
                                    st.rerun()
            else:
                st.info("登録されているイベントはありません。")

        # --- タブ3: 回答結果の確認 ---
        with tab_answers:
            st.subheader("📊 LINEでの回答結果")
            df_answers = get_sheet_data("イベント回答")

            if not df_answers.empty and "イベント名" in df_answers.columns:
                event_list = sorted(df_answers["イベント名"].dropna().unique().tolist())
                selected_event = st.selectbox("確認したいイベントを選択", event_list)

                df_filtered = df_answers[df_answers["イベント名"] == selected_event].copy()

                if not df_filtered.empty:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("総回答数", f"{len(df_filtered)} 件")
                    
                    st.divider()

                    # 集計結果
                    st.markdown("#### 📈 回答の集計")
                    summary = df_filtered["回答"].value_counts().reset_index()
                    summary.columns = ["回答", "人数"]
                    st.dataframe(summary, use_container_width=True)

                    st.markdown("#### 📋 回答明細一覧")
                    display_cols = [c for c in ["氏名", "回答", "コメント", "回答日時"] if c in df_filtered.columns]
                    st.dataframe(df_filtered[display_cols], use_container_width=True)

                    # Excelダウンロード
                    def create_event_excel(df, event_title):
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "回答一覧"

                        ws.merge_cells("A1:D1")
                        title_cell = ws["A1"]
                        title_cell.value = f"回答結果：{event_title}"
                        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
                        title_cell.fill = PatternFill("solid", fgColor="2E7D32")
                        title_cell.alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[1].height = 30

                        headers = list(df.columns)
                        header_fill = PatternFill("solid", fgColor="4CAF50")
                        thin = Side(style="thin", color="CCCCCC")
                        border = Border(left=thin, right=thin, top=thin, bottom=thin)

                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=2, column=col_idx, value=header)
                            cell.font = Font(bold=True, color="FFFFFF", size=11)
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.border = border
                        ws.row_dimensions[2].height = 22

                        for row_idx, row in df.iterrows():
                            for col_idx, value in enumerate(row.values, 1):
                                cell = ws.cell(row=row_idx + 3, column=col_idx, value=str(value))
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.border = border
                                if row_idx % 2 == 0:
                                    cell.fill = PatternFill("solid", fgColor="E8F5E9")
                            ws.row_dimensions[row_idx + 3].height = 20

                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        return output

                    excel_data = create_event_excel(df_filtered[display_cols], selected_event)
                    st.download_button(
                        label="📥 回答結果をExcelでダウンロード",
                        data=excel_data,
                        file_name=f"イベント回答_{selected_event}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.info("このイベントの回答はまだありません。")
            else:
                st.info("まだLINEからのイベント回答データがありません。")

    # ==========================================
    # 勤怠一覧
    # ==========================================
    elif menu == "📊 勤怠一覧":
        st.header("📊 勤怠一覧 Excelダウンロード")

        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "開始日",
                value=datetime.now(JST).date().replace(day=1)
            )
        with col2:
            end_date = st.date_input(
                "終了日",
                value=datetime.now(JST).date()
            )

        if start_date > end_date:
            st.error("終了日は開始日以降にしてください")
            return

        df = get_sheet_data("欠勤連絡")

        if df.empty:
            st.info("データがありません")
        else:
            df["日付_正規化"] = df["日付"].apply(normalize_date)
            start_str = start_date.strftime("%Y/%m/%d")
            end_str = end_date.strftime("%Y/%m/%d")
            df_filtered = df[
                (df["日付_正規化"] >= start_str) &
                (df["日付_正規化"] <= end_str)
            ].copy()

            departments = ["すべて"] + sorted(df_filtered["部署"].dropna().unique().tolist())
            selected_dept = st.selectbox("部署で絞り込み", departments)
            if selected_dept != "すべて":
                df_filtered = df_filtered[df_filtered["部署"] == selected_dept]

            names = ["すべて"] + sorted(df_filtered["氏名"].dropna().unique().tolist())
            selected_name = st.selectbox("氏名で絞り込み", names)
            if selected_name != "すべて":
                df_filtered = df_filtered[df_filtered["氏名"] == selected_name]

            display_cols = ["日付", "氏名", "部署", "種別", "開始時刻", "終了時刻", "備考", "登録時刻"]
            df_display = df_filtered[[col for col in display_cols if col in df_filtered.columns]].copy()
            df_display = df_display.sort_values("日付").reset_index(drop=True)

            st.subheader(f"📋 該当件数：{len(df_display)}件")
            st.dataframe(df_display, use_container_width=True)

            def create_excel(df):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "勤怠一覧"

                ws.merge_cells("A1:H1")
                title_cell = ws["A1"]
                title_cell.value = f"勤怠一覧 {start_date.strftime('%Y/%m/%d')} ～ {end_date.strftime('%Y/%m/%d')}"
                title_cell.font = Font(bold=True, size=14, color="FFFFFF")
                title_cell.fill = PatternFill("solid", fgColor="1565C0")
                title_cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 30

                headers = list(df.columns)
                header_fill = PatternFill("solid", fgColor="0288D1")
                thin = Side(style="thin", color="CCCCCC")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)

                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_idx, value=header)
                    cell.font = Font(bold=True, color="FFFFFF", size=11)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border
                ws.row_dimensions[2].height = 22

                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row.values, 1):
                        cell = ws.cell(row=row_idx + 3, column=col_idx, value=value)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = border
                        if row_idx % 2 == 0:
                            cell.fill = PatternFill("solid", fgColor="E3F2FD")
                    ws.row_dimensions[row_idx + 3].height = 20

                column_widths = {
                    "日付": 14,
                    "氏名": 14,
                    "部署": 14,
                    "種別": 12,
                    "開始時刻": 12,
                    "終了時刻": 12,
                    "備考": 24,
                    "登録時刻": 12,
                }
                for col_idx, header in enumerate(headers, 1):
                    ws.column_dimensions[
                        openpyxl.utils.get_column_letter(col_idx)
                    ].width = column_widths.get(header, 14)

                output = BytesIO()
                wb.save(output)
                output.seek(0)
                return output

            if not df_display.empty:
                excel_data = create_excel(df_display)
                file_name = f"勤怠一覧_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    label="📥 Excelダウンロード",
                    data=excel_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("該当するデータがありません")

    # ==========================================
    # 🆘 安否確認一覧
    # ==========================================
    elif menu == "🆘 安否確認一覧":
        st.header("🆘 安否確認一覧")

        # 日付フィルター
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "開始日",
                value=datetime.now(JST).date()
            )
        with col2:
            end_date = st.date_input(
                "終了日",
                value=datetime.now(JST).date()
            )

        if start_date > end_date:
            st.error("終了日は開始日以降にしてください")
            return

        # データ取得
        df = get_sheet_data("安否確認")

        if df.empty:
            st.info("安否確認のデータがありません")
            return

        # 日付フィルター
        # 列名は「日時」！！
        if "日時" in df.columns:
            df["日付_正規化"] = df["日時"].apply(
                lambda x: normalize_date(str(x).split(" ")[0]) if x else ""
            )
            start_str = start_date.strftime("%Y/%m/%d")
            end_str = end_date.strftime("%Y/%m/%d")
            df_filtered = df[
                (df["日付_正規化"] >= start_str) &
                (df["日付_正規化"] <= end_str)
            ].copy()
        else:
            df_filtered = df.copy()

        # 件数表示
        total = len(df_filtered)
        safe = len(df_filtered[df_filtered["ステータス"] == "無事です"]) \
            if "ステータス" in df_filtered.columns else total

        # サマリーカード
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="✅ 回答件数", value=f"{total} 件")
        with col2:
            st.metric(label="🟢 無事です", value=f"{safe} 件")

        st.divider()

        # 一覧表示
        st.subheader(f"📋 回答一覧（{total}件）")

        if not df_filtered.empty:
            # 表示列を選択
            # 列名は「日時」「LINE表示名」「ステータス」！！
            display_cols = []
            for col in ["日時", "LINE表示名", "ステータス"]:
                if col in df_filtered.columns:
                    display_cols.append(col)

            df_display = df_filtered[display_cols].copy()

            # 日時で降順ソート
            if "日時" in df_display.columns:
                df_display = df_display.sort_values(
                    "日時", ascending=False
                ).reset_index(drop=True)

            # カード表示
            df_filtered_reset = df_filtered.copy()
            df_filtered_reset.reset_index(drop=False, inplace=True)

            for _, row in df_filtered_reset.iterrows():
                col1, col2 = st.columns([4, 1])
                with col1:
                    status = row.get("ステータス", "無事です")
                    icon = "✅" if status == "無事です" else "⚠️"
                    # LINE表示名を使う！！
                    st.success(
                        f"{icon} {row.get('LINE表示名', '')} / "
                        f"{status} / "
                        f"🕐 {row.get('日時', '')}"
                    )
                with col2:
                    if st.button("🗑️ 削除", key=f"anpi_{row['index']}"):
                        if delete_row("安否確認", row["index"]):
                            st.success("削除しました！")
                            st.cache_resource.clear()
                            st.rerun()

            st.divider()

            # Excelダウンロード
            def create_anpi_excel(df):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "安否確認一覧"

                # タイトル
                ws.merge_cells("A1:C1")
                title_cell = ws["A1"]
                title_cell.value = (
                    f"安否確認一覧 "
                    f"{start_date.strftime('%Y/%m/%d')} ～ "
                    f"{end_date.strftime('%Y/%m/%d')}"
                )
                title_cell.font = Font(bold=True, size=14, color="FFFFFF")
                title_cell.fill = PatternFill("solid", fgColor="C62828")
                title_cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 30

                # ヘッダー
                headers = list(df.columns)
                header_fill = PatternFill("solid", fgColor="EF5350")
                thin = Side(style="thin", color="CCCCCC")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)

                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=2, column=col_idx, value=header)
                    cell.font = Font(bold=True, color="FFFFFF", size=11)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border
                ws.row_dimensions[2].height = 22

                # データ
                for row_idx, row in df.iterrows():
                    for col_idx, value in enumerate(row.values, 1):
                        cell = ws.cell(row=row_idx + 3, column=col_idx, value=value)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = border
                        if row_idx % 2 == 0:
                            cell.fill = PatternFill("solid", fgColor="FFEBEE")
                    ws.row_dimensions[row_idx + 3].height = 20

                # 列幅
                column_widths = {
                    "日時": 20,
                    "LINE表示名": 16,
                    "ステータス": 14,
                }
                for col_idx, header in enumerate(headers, 1):
                    ws.column_dimensions[
                        openpyxl.utils.get_column_letter(col_idx)
                    ].width = column_widths.get(header, 16)

                output = BytesIO()
                wb.save(output)
                output.seek(0)
                return output

            excel_data = create_anpi_excel(df_display)
            file_name = (
                f"安否確認一覧_"
                f"{start_date.strftime('%Y%m%d')}_"
                f"{end_date.strftime('%Y%m%d')}.xlsx"
            )
            st.download_button(
                label="📥 Excelダウンロード",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("該当するデータがありません")


if __name__ == "__main__":
    main()
