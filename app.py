"""
個人記帳系統 - Streamlit + Pandas + Google Sheets + yfinance
第三版：新增 📈 帳戶資產走勢圖、📊 圖表分析頁、💰 每月預算、
        📋 明細篩選/搜尋、直接修改/刪除舊記錄

功能：
1. 左側 Sidebar 表單輸入（支出/收入），或轉帳雙帳戶欄位
2. 若幣別不是 TWD，自動用 yfinance 抓當天歷史匯率，並把換算後的 TWD 金額存起來
3. 週末 / 假日防呆：抓不到當天匯率時，自動往前找最近一個交易日的收盤價
4. 轉帳（含 ATM 領現金、換匯、現金存入銀行）用「一出一入」兩列記錄
5. 訂閱/定期支出可設定週期與下次到期日，總覽頁會列出到期提醒，一鍵記一筆
6. 帳戶資產頁：可設定各帳戶初始餘額，畫出每個帳戶隨時間累積的資產走勢
7. 圖表分析頁：每月支出/收入長條圖、分類圓餅圖、專案花費統計
8. 預算頁：設定每月總預算與分類預算，顯示進度條與超支警示
9. 明細管理頁：依日期/帳戶/分類/關鍵字篩選，可直接編輯或刪除舊記錄
10. 所有資料都存在 Google Sheets，換裝置也能繼續用
"""

import calendar
import uuid
from datetime import date, datetime, timedelta

import altair as alt
import gspread
import pandas as pd
import streamlit as st
import yfinance as yf
from google.oauth2.service_account import Credentials

# ========== 基本設定（請依照你自己的 Google Sheet 修改）==========
st.set_page_config(page_title="我的記帳本", page_icon="💰", layout="wide")

SHEET_NAME = "emily記帳本本資料"          # 你的 Google Sheet 檔案名稱
WORKSHEET_NAME = "工作表1"                # 記帳明細分頁（tab）名稱
RECURRING_WORKSHEET_NAME = "訂閱清單"     # 訂閱/定期支出設定分頁（會自動建立）
BUDGET_WORKSHEET_NAME = "預算設定"        # 每月預算分頁（會自動建立）
ACCOUNT_SETTINGS_WORKSHEET_NAME = "帳戶設定"  # 各帳戶初始餘額（會自動建立）

ACCOUNT_OPTIONS = ["郵局", "銀行", "中國信託", "證券戶", "現金", "美國銀行"]
TYPE_OPTIONS = ["支出", "收入", "轉帳"]
CURRENCY_OPTIONS = ["TWD", "USD", "EUR"]
PROJECT_OPTIONS = ["常規", "美國2026旅行", "美國2026打工度假"]
FREQUENCY_OPTIONS = ["每週", "每月", "每季", "每年"]

# 類別下拉選單選項；選「其他（自訂）」時可自行輸入文字
CATEGORY_OPTIONS = ["飲食", "交通", "住宿", "日用品", "娛樂", "學習", "訂閱", "醫療", "其他（自訂）"]
CUSTOM_CATEGORY_KEY = "其他（自訂）"

# 各類別的子類別選項；沒列在這裡的類別，子類別為自由輸入
SUBCATEGORY_MAP = {
    "飲食": ["早餐", "午餐", "晚餐", "零食", "點心", "飲料", "超市日常採買"],
}

TOTAL_BUDGET_KEY = "總預算"  # 預算設定分頁中，這個分類名稱代表「當月全部支出」的預算

# Item = 商家/品項、Note = 備註（新欄位加在最後，ensure_worksheet 會自動幫舊表補表頭）
TRANSACTIONS_HEADER = [
    "Date", "Type", "Account", "Category", "Sub_category",
    "Amount", "Currency", "Project", "TransferID", "Amount_TWD",
    "Item", "Note",
]
RECURRING_HEADER = [
    "Name", "Amount", "Currency", "Account", "Category", "Sub_Category",
    "Frequency", "NextDueDate", "Project", "Active",
]
BUDGET_HEADER = ["Category", "MonthlyBudget_TWD"]
ACCOUNT_SETTINGS_HEADER = ["Account", "InitialBalance_TWD"]


# ========== Google Sheets 連線 ==========
@st.cache_resource
def get_spreadsheet():
    """建立並快取 Google Sheets 連線，回傳整個試算表物件。"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME)


@st.cache_resource
def ensure_worksheet(_spreadsheet, name, header):
    """
    取得指定分頁，若不存在就自動建立並寫入表頭；
    若分頁已存在但表頭比預期少幾欄（例如舊資料只有前 8 欄），
    自動把表頭補到完整長度，不動既有欄位順序。
    """
    header = list(header)
    try:
        ws = _spreadsheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = _spreadsheet.add_worksheet(title=name, rows=1000, cols=len(header) + 5)
        ws.append_row(header)
        return ws

    existing = ws.row_values(1)
    if not existing:
        ws.append_row(header)
    elif len(existing) < len(header) and header[: len(existing)] == existing:
        ws.update(values=[header], range_name="A1")
    return ws


def overwrite_worksheet(ws, header, rows):
    """整張小型設定表重寫（預算、帳戶初始餘額這類列數很少的分頁用）。"""
    ws.clear()
    values = [list(header)] + [list(r) for r in rows]
    ws.update(values=values, range_name="A1")


# ========== 核心：匯率抓取（含週末/假日防呆） ==========
def get_exchange_rate(target_date: date, currency: str):
    """
    抓取指定日期「currency -> TWD」的歷史匯率。

    防呆邏輯：
    不是用「今天是星期幾」去判斷，而是直接往前抓 7 天的歷史資料當緩衝，
    然後篩選出「小於等於記帳日」的所有資料，取最後一筆(也就是離記帳日
    最近、且有收盤價的那個交易日)。這樣不只涵蓋週六日，連國定假日、
    連續假期都能自動往前抓到最近的有效匯率，不會抓到 NaN。

    回傳: (匯率 float, 實際採用的交易日 date)
    """
    if currency == "TWD":
        return 1.0, target_date

    ticker_symbol = f"{currency}TWD=X"  # 例如 USDTWD=X, EURTWD=X

    start_date = target_date - timedelta(days=7)
    end_date = target_date + timedelta(days=1)  # yfinance 的 end 為不包含當天，故 +1

    try:
        hist = yf.Ticker(ticker_symbol).history(
            start=start_date.strftime("%Y-%m-%d"),
            end=end_date.strftime("%Y-%m-%d"),
        )
    except Exception as e:
        raise RuntimeError(f"抓取匯率時發生連線錯誤：{e}")

    if hist.empty:
        raise ValueError(
            f"抓不到 {currency}/TWD 在 {target_date} 前後的匯率資料，"
            "請確認幣別代碼是否正確，或稍後再試一次。"
        )

    # 統一拿掉時區資訊，避免日期比較時出錯
    hist.index = hist.index.tz_localize(None)

    # 只保留「小於等於記帳日」的資料
    valid = hist[hist.index.date <= target_date]

    if valid.empty:
        # 極端狀況：往前 7 天都抓不到，退而求其次用抓到的最早一筆頂著用
        valid = hist

    last_row = valid.iloc[-1]
    actual_date = valid.index[-1].date()
    rate = float(last_row["Close"])

    return rate, actual_date


# ========== 日期/週期輔助函式（給訂閱用） ==========
def add_months(d: date, months: int) -> date:
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def advance_date(d: date, frequency: str) -> date:
    if frequency == "每週":
        return d + timedelta(days=7)
    if frequency == "每月":
        return add_months(d, 1)
    if frequency == "每季":
        return add_months(d, 3)
    if frequency == "每年":
        return add_months(d, 12)
    return d


# ========== 寫入單筆交易（支出/收入/轉帳的其中一腳） ==========
def append_transaction(ws, date_, type_, account, category, sub_category,
                        amount, currency, project, transfer_id="", amount_twd=None,
                        item="", note=""):
    if amount_twd is None:
        rate, _ = get_exchange_rate(date_, currency)
        amount_twd = round(amount * rate, 2)
    row = [
        date_.strftime("%Y-%m-%d"), type_, account, category, sub_category,
        amount, currency, project, transfer_id, amount_twd,
        item, note,
    ]
    ws.append_row(row)
    return amount_twd


# ========== 左側 Sidebar：新增一筆（支出/收入/轉帳） ==========
def render_sidebar_form():
    st.sidebar.header("📝 新增一筆")
    entry_type = st.sidebar.selectbox("類型 Type", TYPE_OPTIONS, key="entry_type")

    # 類別放在表單外，改變類別時子類別選項才會即時跟著換
    category = ""
    if entry_type != "轉帳":
        category = st.sidebar.selectbox("類別 Category", CATEGORY_OPTIONS, key="entry_category")
        if category == CUSTOM_CATEGORY_KEY:
            category = st.sidebar.text_input("自訂類別名稱", key="entry_category_custom").strip()

    with st.sidebar.form(key="entry_form", clear_on_submit=True):
        input_date = st.date_input("日期 Date", value=date.today())

        if entry_type == "轉帳":
            from_account = st.selectbox("轉出帳戶 From", ACCOUNT_OPTIONS, key="from_acc")
            to_account = st.selectbox("轉入帳戶 To", ACCOUNT_OPTIONS, key="to_acc")
            from_amount = st.number_input("轉出金額 From Amount", min_value=0.0, step=1.0, format="%.2f")
            from_currency = st.selectbox("轉出幣別 From Currency", CURRENCY_OPTIONS, key="from_cur")
            to_amount = st.number_input(
                "轉入金額 To Amount（實際到手金額，同幣別請填一樣）",
                min_value=0.0, step=1.0, format="%.2f",
            )
            to_currency = st.selectbox("轉入幣別 To Currency", CURRENCY_OPTIONS, key="to_cur")
            note = st.text_input("備註 Note（例如：ATM領現金、換匯）")
            project = st.selectbox("專案 Project", PROJECT_OPTIONS, key="transfer_project")
        else:
            account = st.selectbox("帳戶 Account", ACCOUNT_OPTIONS)
            sub_options = SUBCATEGORY_MAP.get(category)
            if sub_options:
                sub_category = st.selectbox("子類別 Sub_Category", sub_options)
            else:
                sub_category = st.text_input("子類別 Sub_Category（可留空）")
            item = st.text_input("商家/品項 Item（例如：全聯-鮮奶）")
            amount = st.number_input("金額 Amount", min_value=0.0, step=1.0, format="%.2f")
            currency = st.selectbox("幣別 Currency", CURRENCY_OPTIONS)
            project = st.selectbox("專案 Project", PROJECT_OPTIONS)
            note = st.text_input("備註 Note（可留空）")

        submitted = st.form_submit_button("送出 ✅")

    if submitted:
        if entry_type == "轉帳":
            handle_submit_transfer(
                input_date, from_account, to_account,
                from_amount, from_currency, to_amount, to_currency,
                note, project,
            )
        else:
            handle_submit_normal(
                input_date, entry_type, account, category,
                sub_category, amount, currency, project, item, note,
            )


def handle_submit_normal(input_date, input_type, account, category,
                          sub_category, amount, currency, project,
                          item="", note=""):
    if not category.strip():
        st.sidebar.error("請先選擇或輸入類別。")
        return
    if amount <= 0:
        st.sidebar.error("金額必須大於 0，請重新輸入。")
        return

    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, WORKSHEET_NAME, tuple(TRANSACTIONS_HEADER))

    with st.spinner("處理中：正在確認匯率並寫入 Google Sheets..."):
        try:
            rate, rate_date = get_exchange_rate(input_date, currency)
        except Exception as e:
            st.sidebar.error(f"匯率抓取失敗：{e}")
            return

        amount_twd = round(amount * rate, 2)

        try:
            append_transaction(
                ws, input_date, input_type, account, category, sub_category,
                amount, currency, project, "", amount_twd, item, note,
            )
        except Exception as e:
            st.sidebar.error(f"寫入 Google Sheets 失敗：{e}")
            return

    st.sidebar.success("✅ 已成功記錄一筆帳！")

    if currency != "TWD":
        note = f"採用 {rate_date} 收盤匯率 {rate:.4f}，換算約 {amount_twd} TWD。"
        if rate_date != input_date:
            note += f"\n⚠️ {input_date} 為非交易日（週末或假日），已自動往前取用最近一個交易日的匯率。"
        st.sidebar.info(note)


def handle_submit_transfer(input_date, from_account, to_account,
                            from_amount, from_currency, to_amount, to_currency,
                            note, project):
    if from_amount <= 0 or to_amount <= 0:
        st.sidebar.error("轉出/轉入金額都必須大於 0。")
        return
    if from_account == to_account:
        st.sidebar.error("轉出帳戶與轉入帳戶不能相同。")
        return

    transfer_id = uuid.uuid4().hex[:8]
    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, WORKSHEET_NAME, tuple(TRANSACTIONS_HEADER))

    with st.spinner("處理中：正在確認匯率並寫入 Google Sheets..."):
        try:
            append_transaction(
                ws, input_date, "轉帳", from_account, "轉帳", note,
                -from_amount, from_currency, project, transfer_id,
            )
            append_transaction(
                ws, input_date, "轉帳", to_account, "轉帳", note,
                to_amount, to_currency, project, transfer_id,
            )
        except Exception as e:
            st.sidebar.error(f"轉帳記錄失敗：{e}")
            return

    st.sidebar.success(f"✅ 已記錄轉帳：{from_account} → {to_account}")
    if from_currency != to_currency:
        implied_rate = to_amount / from_amount
        st.sidebar.info(f"這筆換匯的實質匯率約為 1 {from_currency} = {implied_rate:.4f} {to_currency}")


# ========== 左側 Sidebar：訂閱 / 定期支出管理 ==========
def render_recurring_manager():
    with st.sidebar.expander("🔁 訂閱 / 定期支出管理"):
        with st.form(key="recurring_form", clear_on_submit=True):
            name = st.text_input("訂閱名稱 Name（例如：Netflix）")
            account = st.selectbox("帳戶 Account", ACCOUNT_OPTIONS, key="rec_acc")
            category = st.text_input("類別 Category", key="rec_cat")
            sub_category = st.text_input("店家/品項 Sub_Category", key="rec_sub")
            amount = st.number_input("金額 Amount", min_value=0.0, step=1.0, format="%.2f", key="rec_amt")
            currency = st.selectbox("幣別 Currency", CURRENCY_OPTIONS, key="rec_cur")
            frequency = st.selectbox("週期 Frequency", FREQUENCY_OPTIONS, key="rec_freq")
            start_date = st.date_input("下次扣款日 Next Due Date", value=date.today(), key="rec_start")
            project = st.selectbox("專案 Project", PROJECT_OPTIONS, key="rec_project")
            submitted = st.form_submit_button("新增訂閱 ✅")

        if submitted:
            handle_add_recurring(name, account, category, sub_category, amount, currency, frequency, start_date, project)


def handle_add_recurring(name, account, category, sub_category, amount, currency, frequency, start_date, project):
    if not name.strip():
        st.sidebar.error("請輸入訂閱名稱。")
        return
    if amount <= 0:
        st.sidebar.error("金額必須大於 0。")
        return

    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, RECURRING_WORKSHEET_NAME, tuple(RECURRING_HEADER))
    row = [
        name, amount, currency, account, category, sub_category,
        frequency, start_date.strftime("%Y-%m-%d"), project, "是",
    ]
    ws.append_row(row)
    st.sidebar.success(f"✅ 已新增訂閱：{name}")


# ========== 總覽頁：訂閱提醒 ==========
def render_recurring_reminders():
    st.subheader("🔔 訂閱 / 定期支出提醒")

    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, RECURRING_WORKSHEET_NAME, tuple(RECURRING_HEADER))
    records = ws.get_all_records()

    if not records:
        st.caption("目前沒有設定任何訂閱項目，可以在左側「訂閱 / 定期支出管理」新增。")
        return

    today = date.today()
    entries = []
    for idx, r in enumerate(records):
        if str(r.get("Active", "是")) not in ("是", "TRUE", "True", "1"):
            continue
        try:
            next_due = datetime.strptime(str(r["NextDueDate"]), "%Y-%m-%d").date()
        except Exception:
            continue
        entries.append((idx, r, next_due, (next_due - today).days))

    if not entries:
        st.caption("目前沒有啟用中的訂閱項目。")
        return

    entries.sort(key=lambda x: x[2])

    for idx, r, next_due, days_left in entries:
        if days_left < 0:
            status = f"⚠️ 已逾期 {abs(days_left)} 天"
        elif days_left <= 7:
            status = f"⏰ 還有 {days_left} 天到期"
        else:
            status = f"{next_due} 到期"

        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(f"**{r['Name']}** - {r['Amount']} {r['Currency']} - {r['Account']} - {status}")
        with col2:
            if days_left <= 7 and st.button("✅ 記一筆", key=f"rec_pay_{idx}"):
                record_recurring_payment(idx, r, next_due)
                st.rerun()


def record_recurring_payment(idx, rule, due_date):
    spreadsheet = get_spreadsheet()
    tx_ws = ensure_worksheet(spreadsheet, WORKSHEET_NAME, tuple(TRANSACTIONS_HEADER))
    rec_ws = ensure_worksheet(spreadsheet, RECURRING_WORKSHEET_NAME, tuple(RECURRING_HEADER))

    try:
        append_transaction(
            tx_ws, due_date, "支出", rule["Account"], rule["Category"], rule["Sub_Category"],
            float(rule["Amount"]), rule["Currency"], rule.get("Project", ""),
        )
    except Exception as e:
        st.error(f"記錄訂閱扣款失敗：{e}")
        return

    next_due = advance_date(due_date, rule["Frequency"])
    sheet_row = idx + 2
    col = RECURRING_HEADER.index("NextDueDate") + 1
    rec_ws.update_cell(sheet_row, col, next_due.strftime("%Y-%m-%d"))
    st.success(f"✅ 已記錄「{rule['Name']}」本期扣款，下次到期日更新為 {next_due}")


# ========== 總覽頁：各帳戶總覽 ==========
def render_account_overview(df):
    valid = df.dropna(subset=["Amount_TWD"])
    rows = []
    for acct in ACCOUNT_OPTIONS:
        sub = valid[valid["Account"] == acct]
        if sub.empty:
            continue
        total_expense = sub[sub["Type"] == "支出"]["Amount_TWD"].sum()
        total_income = sub[sub["Type"] == "收入"]["Amount_TWD"].sum()
        transfer_net = sub[sub["Type"] == "轉帳"]["Amount_TWD"].sum()
        net_change = total_income - total_expense + transfer_net
        rows.append({
            "帳戶": acct,
            "總支出(TWD)": round(total_expense, 2),
            "總收入(TWD)": round(total_income, 2),
            "轉帳淨額(TWD)": round(transfer_net, 2),
            "淨變動(TWD)": round(net_change, 2),
        })

    if rows:
        st.dataframe(pd.DataFrame(rows), width="stretch")
    else:
        st.caption("目前沒有可統計的帳戶資料。")


# ========== 總覽頁：同分類同店家本月消費 2 次以上 ==========
def render_merchant_aggregation(df):
    exp_df = df[df["Type"] == "支出"].copy()
    # 新資料的商家/品項在 Item 欄；舊資料在 Sub_category 欄，取有值的那個
    items = exp_df["Item"].astype(str).str.strip()
    exp_df["商家/品項"] = items.where(items != "", exp_df["Sub_category"].astype(str).str.strip())
    exp_df = exp_df[exp_df["商家/品項"] != ""]
    if exp_df.empty:
        st.caption("目前沒有足夠的支出明細可供統計。")
        return

    months = sorted(exp_df["YearMonth"].dropna().unique(), reverse=True)
    if not months:
        st.caption("目前沒有足夠的支出明細可供統計。")
        return

    selected_month = st.selectbox("選擇月份", months, index=0, key="merchant_month")
    month_df = exp_df[exp_df["YearMonth"] == selected_month]

    grouped = (
        month_df.groupby(["Category", "商家/品項"])
        .agg(次數=("Amount_TWD", "count"), 台幣總額=("Amount_TWD", "sum"))
        .reset_index()
    )
    grouped = grouped[grouped["次數"] >= 2].sort_values("台幣總額", ascending=False)

    if grouped.empty:
        st.caption(f"{selected_month} 沒有同分類同店家消費 2 次以上的紀錄。")
    else:
        st.dataframe(grouped, width="stretch")


# ========== 補齊舊資料缺少的台幣換算金額 ==========
def backfill_amount_twd(ws, df):
    col = TRANSACTIONS_HEADER.index("Amount_TWD") + 1
    updated = 0
    for i, row in df.iterrows():
        if pd.isna(row["Amount_TWD"]):
            try:
                d = datetime.strptime(str(row["Date"]), "%Y-%m-%d").date()
                rate, _ = get_exchange_rate(d, row["Currency"])
                twd = round(float(row["Amount"]) * rate, 2)
                ws.update_cell(i + 2, col, twd)
                updated += 1
            except Exception:
                continue
    st.success(f"已補齊 {updated} 筆歷史資料的台幣換算金額。")


# ========== 共用：帶正負號的台幣金額（支出為負、收入為正、轉帳照原本正負） ==========
def add_signed_twd(df):
    def _sign(row):
        twd = row["Amount_TWD"]
        if pd.isna(twd):
            return twd
        if row["Type"] == "支出":
            return -abs(twd)
        if row["Type"] == "收入":
            return abs(twd)
        return twd  # 轉帳寫入時金額本身就帶正負號

    df = df.copy()
    df["Signed_TWD"] = df.apply(_sign, axis=1)
    return df


# ========== 📈 帳戶資產頁 ==========
def load_initial_balances():
    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, ACCOUNT_SETTINGS_WORKSHEET_NAME, tuple(ACCOUNT_SETTINGS_HEADER))
    balances = {}
    for r in ws.get_all_records():
        acct = str(r.get("Account", "")).strip()
        if not acct:
            continue
        try:
            balances[acct] = float(r.get("InitialBalance_TWD", 0) or 0)
        except (TypeError, ValueError):
            balances[acct] = 0.0
    return ws, balances


def render_assets_tab(df):
    st.subheader("📈 各帳戶資產走勢")
    st.caption(
        "資產 = 初始餘額 + 記帳以來的所有變動（外幣以記帳當天匯率換算成 TWD）。"
        "想讓數字貼近銀行實際餘額，請先在下方設定各帳戶的初始餘額。"
    )

    ws, balances = load_initial_balances()

    with st.expander("⚙️ 設定各帳戶初始餘額（開始記帳那天的餘額，TWD）"):
        edit_df = pd.DataFrame(
            [{"帳戶": a, "初始餘額(TWD)": balances.get(a, 0.0)} for a in ACCOUNT_OPTIONS]
        )
        edited = st.data_editor(
            edit_df,
            hide_index=True,
            width="stretch",
            column_config={
                "帳戶": st.column_config.TextColumn(disabled=True),
                "初始餘額(TWD)": st.column_config.NumberColumn(format="%.2f"),
            },
            key="init_balance_editor",
        )
        if st.button("💾 儲存初始餘額"):
            rows = [
                [str(r["帳戶"]), float(r["初始餘額(TWD)"] or 0)]
                for _, r in edited.iterrows()
            ]
            try:
                overwrite_worksheet(ws, ACCOUNT_SETTINGS_HEADER, rows)
            except Exception as e:
                st.error(f"儲存失敗：{e}")
                return
            st.success("✅ 初始餘額已儲存！")
            st.rerun()

    if df.empty:
        st.info("還沒有記帳資料，先從左側記幾筆帳，這裡就會長出走勢圖。")
        return

    valid = add_signed_twd(df).dropna(subset=["Signed_TWD", "Date_parsed"])
    if valid.empty:
        st.caption("目前沒有可統計的資料（可能缺少台幣換算金額）。")
        return

    # 每個帳戶每天的變動加總 -> 累積 -> 加上初始餘額
    pivot = valid.pivot_table(
        index="Date_parsed", columns="Account", values="Signed_TWD", aggfunc="sum"
    ).sort_index()
    cumulative = pivot.fillna(0).cumsum()
    for acct in cumulative.columns:
        cumulative[acct] = cumulative[acct] + balances.get(acct, 0.0)

    # 有設定初始餘額但還沒有任何交易的帳戶，也要以水平線的方式納入資產
    for acct, bal in balances.items():
        if acct not in cumulative.columns and bal != 0:
            cumulative[acct] = bal

    # 目前各帳戶估計餘額
    latest = cumulative.iloc[-1]
    total_assets = float(latest.sum())

    metric_cols = st.columns(len(latest) + 1)
    metric_cols[0].metric("💎 總資產(TWD)", f"{total_assets:,.0f}")
    for i, (acct, val) in enumerate(latest.items(), start=1):
        metric_cols[i].metric(acct, f"{val:,.0f}")

    chart_df = cumulative.reset_index().melt(
        id_vars="Date_parsed", var_name="帳戶", value_name="資產(TWD)"
    )
    line = (
        alt.Chart(chart_df)
        .mark_line(point=True)
        .encode(
            x=alt.X("Date_parsed:T", title="日期"),
            y=alt.Y("資產(TWD):Q", title="資產 (TWD)"),
            color=alt.Color("帳戶:N", title="帳戶"),
            tooltip=[
                alt.Tooltip("Date_parsed:T", title="日期"),
                alt.Tooltip("帳戶:N"),
                alt.Tooltip("資產(TWD):Q", format=",.2f"),
            ],
        )
        .properties(height=420)
    )
    st.altair_chart(line, width="stretch")

    # 總資產組成圓餅圖（各帳戶目前餘額的占比）
    st.subheader("💎 總資產組成（各帳戶占比）")
    pie_data = latest.reset_index()
    pie_data.columns = ["帳戶", "餘額(TWD)"]
    negative = pie_data[pie_data["餘額(TWD)"] <= 0]
    pie_data = pie_data[pie_data["餘額(TWD)"] > 0]
    if pie_data.empty:
        st.caption("目前沒有正餘額的帳戶可以畫占比。")
    else:
        pie_data["占比"] = pie_data["餘額(TWD)"] / pie_data["餘額(TWD)"].sum()
        col1, col2 = st.columns([3, 2])
        with col1:
            asset_pie = (
                alt.Chart(pie_data)
                .mark_arc(innerRadius=70)
                .encode(
                    theta=alt.Theta("餘額(TWD):Q"),
                    color=alt.Color("帳戶:N", title="帳戶"),
                    tooltip=[
                        alt.Tooltip("帳戶:N"),
                        alt.Tooltip("餘額(TWD):Q", format=",.0f"),
                        alt.Tooltip("占比:Q", format=".1%"),
                    ],
                )
                .properties(height=340)
            )
            st.altair_chart(asset_pie, width="stretch")
        with col2:
            show = pie_data.copy().sort_values("餘額(TWD)", ascending=False)
            show["餘額(TWD)"] = show["餘額(TWD)"].round(0)
            show["占比"] = (show["占比"] * 100).round(1).astype(str) + "%"
            st.dataframe(show, width="stretch", hide_index=True)
        if not negative.empty:
            names = "、".join(negative["帳戶"].astype(str))
            st.caption(f"⚠️ {names} 目前餘額為 0 或負數，未納入圓餅圖（總資產金額仍有計入）。")

    st.divider()
    render_single_account_view(valid, balances)


# ========== 📈 單一帳戶檢視 ==========
def render_single_account_view(valid, balances):
    st.subheader("🔍 單一帳戶檢視")
    st.caption("選一個帳戶，看它的餘額走勢和每一筆變動。要記帳時在左側表單選同一個帳戶即可。")

    acct = st.selectbox("選擇帳戶", ACCOUNT_OPTIONS, key="single_account")

    sub = valid[valid["Account"] == acct].sort_values("Date_parsed", kind="stable").copy()
    init = balances.get(acct, 0.0)
    current = init + sub["Signed_TWD"].sum()

    this_month = date.today().strftime("%Y-%m")
    m = sub[sub["YearMonth"] == this_month]
    inflow = m[m["Signed_TWD"] > 0]["Signed_TWD"].sum()
    outflow = m[m["Signed_TWD"] < 0]["Signed_TWD"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("目前餘額(TWD)", f"{current:,.0f}")
    c2.metric("初始餘額(TWD)", f"{init:,.0f}")
    c3.metric(f"{this_month} 流入", f"{inflow:,.0f}")
    c4.metric(f"{this_month} 流出", f"{abs(outflow):,.0f}")

    if sub.empty:
        st.caption(f"「{acct}」目前還沒有任何交易記錄。")
        return

    sub["累積餘額(TWD)"] = init + sub["Signed_TWD"].cumsum()

    balance_line = (
        alt.Chart(sub)
        .mark_line(point=True, interpolate="step-after")
        .encode(
            x=alt.X("Date_parsed:T", title="日期"),
            y=alt.Y("累積餘額(TWD):Q", title="餘額 (TWD)"),
            tooltip=[
                alt.Tooltip("Date_parsed:T", title="日期"),
                alt.Tooltip("Category:N", title="分類"),
                alt.Tooltip("Sub_category:N", title="子類別"),
                alt.Tooltip("Item:N", title="商家/品項"),
                alt.Tooltip("Signed_TWD:Q", title="變動(TWD)", format="+,.0f"),
                alt.Tooltip("累積餘額(TWD):Q", format=",.0f"),
            ],
        )
        .properties(height=300)
    )
    st.altair_chart(balance_line, width="stretch")

    history = sub.sort_values("Date_parsed", ascending=False, kind="stable")[
        ["Date", "Type", "Category", "Sub_category", "Item", "Note", "Amount", "Currency",
         "Signed_TWD", "累積餘額(TWD)"]
    ].rename(columns={
        "Signed_TWD": "變動(TWD)", "Sub_category": "子類別",
        "Item": "商家/品項", "Note": "備註",
    })
    history["變動(TWD)"] = history["變動(TWD)"].round(2)
    history["累積餘額(TWD)"] = history["累積餘額(TWD)"].round(2)
    st.dataframe(history, width="stretch", hide_index=True)


# ========== 📊 圖表分析頁 ==========
def render_charts_tab(df):
    if df.empty:
        st.info("還沒有記帳資料，先從左側記幾筆帳吧！")
        return

    valid = df.dropna(subset=["Amount_TWD"])

    # --- 每月支出 / 收入 ---
    st.subheader("📅 每月支出 / 收入")
    monthly = (
        valid[valid["Type"].isin(["支出", "收入"])]
        .groupby(["YearMonth", "Type"])["Amount_TWD"]
        .sum()
        .reset_index()
        .rename(columns={"Amount_TWD": "金額(TWD)"})
    )
    if monthly.empty:
        st.caption("目前沒有支出/收入資料。")
    else:
        bar = (
            alt.Chart(monthly)
            .mark_bar()
            .encode(
                x=alt.X("YearMonth:N", title="月份"),
                y=alt.Y("金額(TWD):Q", title="金額 (TWD)"),
                color=alt.Color(
                    "Type:N", title="類型",
                    scale=alt.Scale(domain=["支出", "收入"], range=["#e45756", "#54a24b"]),
                ),
                xOffset="Type:N",
                tooltip=["YearMonth", "Type", alt.Tooltip("金額(TWD):Q", format=",.0f")],
            )
            .properties(height=340)
        )
        st.altair_chart(bar, width="stretch")

    # --- 分類圓餅圖 ---
    st.subheader("🥧 支出分類占比")
    exp = valid[valid["Type"] == "支出"].copy()
    if exp.empty:
        st.caption("目前沒有支出資料。")
    else:
        months = ["全部時間"] + sorted(exp["YearMonth"].dropna().unique(), reverse=True)
        pie_month = st.selectbox("選擇月份", months, index=0, key="pie_month")
        pie_df = exp if pie_month == "全部時間" else exp[exp["YearMonth"] == pie_month]

        cat = (
            pie_df.groupby("Category")["Amount_TWD"].sum().reset_index()
            .rename(columns={"Amount_TWD": "金額(TWD)"})
            .sort_values("金額(TWD)", ascending=False)
        )
        if cat.empty:
            st.caption(f"{pie_month} 沒有支出資料。")
        else:
            total = cat["金額(TWD)"].sum()
            cat["占比"] = cat["金額(TWD)"] / total

            col1, col2 = st.columns([3, 2])
            with col1:
                pie = (
                    alt.Chart(cat)
                    .mark_arc(innerRadius=60)
                    .encode(
                        theta=alt.Theta("金額(TWD):Q"),
                        color=alt.Color("Category:N", title="分類"),
                        tooltip=[
                            "Category",
                            alt.Tooltip("金額(TWD):Q", format=",.0f"),
                            alt.Tooltip("占比:Q", format=".1%"),
                        ],
                    )
                    .properties(height=340)
                )
                st.altair_chart(pie, width="stretch")
            with col2:
                show = cat.copy()
                show["金額(TWD)"] = show["金額(TWD)"].round(0)
                show["占比"] = (show["占比"] * 100).round(1).astype(str) + "%"
                st.dataframe(show, width="stretch", hide_index=True)

            # --- 子類別占比（延用上面選的月份範圍） ---
            sub_exp = pie_df[pie_df["Sub_category"].astype(str).str.strip() != ""]
            if not sub_exp.empty:
                st.subheader("🍱 子類別占比")
                cat_choices = sorted(sub_exp["Category"].astype(str).unique())
                picked_cat = st.selectbox("選擇類別", cat_choices, key="subcat_category")
                sub_data = (
                    sub_exp[sub_exp["Category"] == picked_cat]
                    .groupby("Sub_category")["Amount_TWD"].agg(["sum", "count"])
                    .reset_index()
                )
                sub_data.columns = ["子類別", "金額(TWD)", "筆數"]
                sub_data = sub_data.sort_values("金額(TWD)", ascending=False)
                sub_bar = (
                    alt.Chart(sub_data)
                    .mark_bar()
                    .encode(
                        x=alt.X("金額(TWD):Q", title="金額 (TWD)"),
                        y=alt.Y("子類別:N", sort="-x"),
                        color=alt.Color("子類別:N", legend=None),
                        tooltip=["子類別", alt.Tooltip("金額(TWD):Q", format=",.0f"), "筆數"],
                    )
                    .properties(height=60 + 36 * len(sub_data))
                )
                st.altair_chart(sub_bar, width="stretch")

    # --- 專案花費統計 ---
    st.subheader("🧳 各專案花費統計")
    proj = valid[valid["Type"] == "支出"].copy()
    proj["Project"] = proj["Project"].replace("", "常規").fillna("常規")
    proj_sum = (
        proj.groupby("Project")["Amount_TWD"].agg(["sum", "count"]).reset_index()
    )
    proj_sum.columns = ["專案", "總支出(TWD)", "筆數"]
    proj_sum = proj_sum.sort_values("總支出(TWD)", ascending=False)
    if proj_sum.empty:
        st.caption("目前沒有專案支出資料。")
    else:
        proj_bar = (
            alt.Chart(proj_sum)
            .mark_bar()
            .encode(
                x=alt.X("總支出(TWD):Q", title="總支出 (TWD)"),
                y=alt.Y("專案:N", sort="-x"),
                color=alt.Color("專案:N", legend=None),
                tooltip=["專案", alt.Tooltip("總支出(TWD):Q", format=",.0f"), "筆數"],
            )
            .properties(height=60 + 40 * len(proj_sum))
        )
        st.altair_chart(proj_bar, width="stretch")


# ========== 💰 預算頁 ==========
def load_budgets():
    spreadsheet = get_spreadsheet()
    ws = ensure_worksheet(spreadsheet, BUDGET_WORKSHEET_NAME, tuple(BUDGET_HEADER))
    budgets = []
    for r in ws.get_all_records():
        cat = str(r.get("Category", "")).strip()
        if not cat:
            continue
        try:
            amt = float(r.get("MonthlyBudget_TWD", 0) or 0)
        except (TypeError, ValueError):
            continue
        if amt > 0:
            budgets.append({"Category": cat, "MonthlyBudget_TWD": amt})
    return ws, budgets


def render_budget_tab(df):
    st.subheader("💰 每月預算")
    st.caption(
        f"在下方表格設定各分類的每月預算（TWD）。分類名稱填「{TOTAL_BUDGET_KEY}」"
        "的那一列代表當月「全部支出」的總預算。"
    )

    ws, budgets = load_budgets()

    # --- 預算設定 ---
    with st.expander("⚙️ 編輯預算設定", expanded=not budgets):
        base = pd.DataFrame(budgets) if budgets else pd.DataFrame(
            [{"Category": TOTAL_BUDGET_KEY, "MonthlyBudget_TWD": 0.0}]
        )
        base = base.rename(columns={"Category": "分類", "MonthlyBudget_TWD": "每月預算(TWD)"})
        edited = st.data_editor(
            base,
            num_rows="dynamic",
            hide_index=True,
            width="stretch",
            column_config={
                "分類": st.column_config.TextColumn(required=True),
                "每月預算(TWD)": st.column_config.NumberColumn(format="%.0f", min_value=0),
            },
            key="budget_editor",
        )
        if st.button("💾 儲存預算設定"):
            rows = []
            for _, r in edited.iterrows():
                cat = str(r["分類"] or "").strip()
                try:
                    amt = float(r["每月預算(TWD)"] or 0)
                except (TypeError, ValueError):
                    amt = 0.0
                if cat and amt > 0:
                    rows.append([cat, amt])
            try:
                overwrite_worksheet(ws, BUDGET_HEADER, rows)
            except Exception as e:
                st.error(f"儲存失敗：{e}")
                return
            st.success("✅ 預算設定已儲存！")
            st.rerun()

    if not budgets:
        st.info("先在上方設定預算，這裡就會顯示每月進度。")
        return

    # --- 預算進度 ---
    if df.empty:
        st.caption("還沒有記帳資料。")
        return

    valid = df.dropna(subset=["Amount_TWD"])
    exp = valid[valid["Type"] == "支出"]

    current_month = date.today().strftime("%Y-%m")
    months = sorted(set(exp["YearMonth"].dropna()) | {current_month}, reverse=True)
    sel_month = st.selectbox("查看月份", months, index=months.index(current_month), key="budget_month")

    month_exp = exp[exp["YearMonth"] == sel_month]
    total_spent = month_exp["Amount_TWD"].sum()

    st.markdown(f"### {sel_month} 預算進度")
    for b in budgets:
        cat = b["Category"]
        budget_amt = b["MonthlyBudget_TWD"]
        if cat == TOTAL_BUDGET_KEY:
            spent = total_spent
            label = f"🌟 {TOTAL_BUDGET_KEY}（全部支出）"
        else:
            spent = month_exp[month_exp["Category"] == cat]["Amount_TWD"].sum()
            label = f"📌 {cat}"

        ratio = spent / budget_amt if budget_amt > 0 else 0
        remaining = budget_amt - spent

        st.write(f"**{label}**：{spent:,.0f} / {budget_amt:,.0f} TWD（{ratio:.0%}）")
        st.progress(min(ratio, 1.0))
        if ratio > 1.0:
            st.error(f"🚨 已超支 {abs(remaining):,.0f} TWD！")
        elif ratio >= 0.8:
            st.warning(f"⚠️ 快到上限了，剩 {remaining:,.0f} TWD。")
        else:
            st.caption(f"還可以花 {remaining:,.0f} TWD。")


# ========== 📋 明細管理頁（篩選 + 編輯 + 刪除） ==========
def render_detail_tab(df, tx_ws):
    st.subheader("📋 記帳明細（可篩選、直接編輯、刪除）")

    if df.empty:
        st.info("目前還沒有任何記帳資料。")
        return

    # --- 篩選列 ---
    f1, f2, f3, f4 = st.columns([2, 1.5, 1.5, 2])
    with f1:
        min_d = df["Date_parsed"].min()
        max_d = df["Date_parsed"].max()
        default_range = (
            (min_d.date(), max_d.date())
            if pd.notna(min_d) and pd.notna(max_d)
            else (date.today(), date.today())
        )
        date_range = st.date_input("日期區間", value=default_range, key="filter_dates")
    with f2:
        sel_types = st.multiselect("類型", TYPE_OPTIONS, key="filter_types")
    with f3:
        accounts = sorted(df["Account"].astype(str).unique())
        sel_accounts = st.multiselect("帳戶", accounts, key="filter_accounts")
    with f4:
        categories = sorted(c for c in df["Category"].astype(str).unique() if c.strip())
        sel_categories = st.multiselect("分類", categories, key="filter_categories")

    keyword = st.text_input("🔍 關鍵字搜尋（分類 / 子類別 / 商家品項 / 備註 / 專案）", key="filter_keyword")

    filtered = df.copy()
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        filtered = filtered[
            (filtered["Date_parsed"] >= pd.Timestamp(start))
            & (filtered["Date_parsed"] <= pd.Timestamp(end))
        ]
    if sel_types:
        filtered = filtered[filtered["Type"].isin(sel_types)]
    if sel_accounts:
        filtered = filtered[filtered["Account"].astype(str).isin(sel_accounts)]
    if sel_categories:
        filtered = filtered[filtered["Category"].astype(str).isin(sel_categories)]
    if keyword.strip():
        kw = keyword.strip()
        mask = (
            filtered["Category"].astype(str).str.contains(kw, case=False, na=False)
            | filtered["Sub_category"].astype(str).str.contains(kw, case=False, na=False)
            | filtered["Item"].astype(str).str.contains(kw, case=False, na=False)
            | filtered["Note"].astype(str).str.contains(kw, case=False, na=False)
            | filtered["Project"].astype(str).str.contains(kw, case=False, na=False)
        )
        filtered = filtered[mask]

    exp_sum = filtered[filtered["Type"] == "支出"]["Amount_TWD"].sum()
    inc_sum = filtered[filtered["Type"] == "收入"]["Amount_TWD"].sum()
    s1, s2, s3 = st.columns(3)
    s1.metric("筆數", len(filtered))
    s2.metric("支出合計(TWD)", f"{exp_sum:,.0f}")
    s3.metric("收入合計(TWD)", f"{inc_sum:,.0f}")

    if filtered.empty:
        st.caption("沒有符合條件的記錄。")
        return

    st.caption(
        "直接在表格裡修改內容或勾選「刪除」，再按下方「💾 儲存變更」。"
        "改了日期/金額/幣別會自動重算台幣金額。"
        "⚠️ 轉帳是一出一入兩筆（TransferID 相同），刪除或修改時記得兩筆都處理。"
    )

    display_cols = [
        "Date", "Type", "Account", "Category", "Sub_category", "Item", "Note",
        "Amount", "Currency", "Project", "TransferID", "Amount_TWD",
    ]
    editor_df = filtered[display_cols].copy()
    editor_df.insert(0, "刪除", False)
    editor_df["Amount"] = pd.to_numeric(editor_df["Amount"], errors="coerce")

    edited = st.data_editor(
        editor_df,
        hide_index=True,
        width="stretch",
        column_config={
            "刪除": st.column_config.CheckboxColumn(help="勾選後按儲存即從 Google Sheets 刪除"),
            "Date": st.column_config.TextColumn(help="格式 YYYY-MM-DD"),
            "Type": st.column_config.SelectboxColumn(options=TYPE_OPTIONS),
            "Account": st.column_config.SelectboxColumn(options=ACCOUNT_OPTIONS),
            "Sub_category": st.column_config.TextColumn("子類別"),
            "Item": st.column_config.TextColumn("商家/品項"),
            "Note": st.column_config.TextColumn("備註"),
            "Amount": st.column_config.NumberColumn(format="%.2f"),
            "Currency": st.column_config.SelectboxColumn(options=CURRENCY_OPTIONS),
            "Project": st.column_config.SelectboxColumn(options=PROJECT_OPTIONS),
            "TransferID": st.column_config.TextColumn(disabled=True),
            "Amount_TWD": st.column_config.NumberColumn(disabled=True, help="自動計算，不用手動改"),
        },
        key="detail_editor",
    )

    if st.button("💾 儲存變更", type="primary"):
        apply_detail_changes(tx_ws, filtered, edited)


def apply_detail_changes(tx_ws, original, edited):
    """比對編輯前後的差異，逐列更新 Google Sheets；勾選刪除的列從下往上刪。"""
    editable_cols = [
        "Date", "Type", "Account", "Category", "Sub_category", "Item", "Note",
        "Amount", "Currency", "Project",
    ]

    to_delete = []   # sheet 列號
    to_update = []   # (sheet列號, 新的一整列資料)
    errors = []

    edited = edited.set_axis(original.index)  # data_editor 會重設索引，對回原本的 sheet 列

    for idx in original.index:
        sheet_row = idx + 2  # 第 1 列是表頭
        if bool(edited.loc[idx, "刪除"]):
            to_delete.append(sheet_row)
            continue

        changed = False
        new_vals = {}
        for col in editable_cols:
            old_v = original.loc[idx, col]
            new_v = edited.loc[idx, col]
            if col == "Amount":
                old_v = float(pd.to_numeric(old_v, errors="coerce") or 0)
                new_v = float(pd.to_numeric(new_v, errors="coerce") or 0)
                if round(old_v, 2) != round(new_v, 2):
                    changed = True
            else:
                old_v = "" if pd.isna(old_v) else str(old_v)
                new_v = "" if pd.isna(new_v) else str(new_v)
                if old_v != new_v:
                    changed = True
            new_vals[col] = new_v

        if not changed:
            continue

        # 驗證日期格式
        try:
            new_date = datetime.strptime(str(new_vals["Date"]).strip(), "%Y-%m-%d").date()
        except ValueError:
            errors.append(f"第 {sheet_row} 列：日期「{new_vals['Date']}」格式錯誤，需為 YYYY-MM-DD，已跳過。")
            continue

        # 日期 / 金額 / 幣別有變就重算台幣金額（轉帳保留正負號）
        amount = new_vals["Amount"]
        old_amount = float(pd.to_numeric(original.loc[idx, "Amount"], errors="coerce") or 0)
        if old_amount < 0 and amount > 0:
            amount = -amount  # 原本是轉出腳（負數），使用者輸入正數時幫忙補回負號

        old_twd = original.loc[idx, "Amount_TWD"]
        need_recalc = (
            str(original.loc[idx, "Date"]) != new_vals["Date"].strip()
            or round(old_amount, 2) != round(amount, 2)
            or str(original.loc[idx, "Currency"]) != new_vals["Currency"]
            or pd.isna(old_twd)
        )
        if need_recalc:
            try:
                rate, _ = get_exchange_rate(new_date, new_vals["Currency"])
                amount_twd = round(amount * rate, 2)
            except Exception as e:
                errors.append(f"第 {sheet_row} 列：匯率抓取失敗（{e}），已跳過這列。")
                continue
        else:
            amount_twd = float(old_twd)

        transfer_id = original.loc[idx, "TransferID"]
        transfer_id = "" if pd.isna(transfer_id) else str(transfer_id)
        row_values = [
            new_date.strftime("%Y-%m-%d"), new_vals["Type"], new_vals["Account"],
            new_vals["Category"], new_vals["Sub_category"],
            amount, new_vals["Currency"], new_vals["Project"],
            transfer_id, amount_twd,
            new_vals["Item"], new_vals["Note"],
        ]
        to_update.append((sheet_row, row_values))

    if not to_delete and not to_update and not errors:
        st.info("沒有偵測到任何變更。")
        return

    updated = deleted = 0
    with st.spinner("正在寫回 Google Sheets..."):
        for sheet_row, row_values in to_update:
            try:
                tx_ws.update(values=[row_values], range_name=f"A{sheet_row}:L{sheet_row}")
                updated += 1
            except Exception as e:
                errors.append(f"第 {sheet_row} 列更新失敗：{e}")

        # 從下往上刪，避免刪除後列號位移
        for sheet_row in sorted(to_delete, reverse=True):
            try:
                tx_ws.delete_rows(sheet_row)
                deleted += 1
            except Exception as e:
                errors.append(f"第 {sheet_row} 列刪除失敗：{e}")

    if updated or deleted:
        st.success(f"✅ 已更新 {updated} 筆、刪除 {deleted} 筆。")
    for msg in errors:
        st.warning(msg)
    if (updated or deleted) and not errors:
        st.rerun()


# ========== 🏠 總覽頁 ==========
def render_overview_tab(df, tx_ws):
    render_recurring_reminders()

    if df.empty:
        st.info("目前還沒有任何記帳資料，從左側表單開始記第一筆帳吧！")
        return

    missing_twd = int(df["Amount_TWD"].isna().sum())
    if missing_twd > 0:
        st.warning(f"有 {missing_twd} 筆舊資料還沒有台幣換算金額（Amount_TWD），統計會先忽略這些筆。")
        if st.button("🔄 補齊缺少的台幣換算金額"):
            backfill_amount_twd(tx_ws, df)
            st.rerun()

    st.subheader("📒 各帳戶總覽（記帳以來的變動，非銀行實際餘額）")
    render_account_overview(df)

    st.subheader("🔁 本月同分類同店家消費 2 次以上")
    render_merchant_aggregation(df)


# ========== 主頁面 ==========
def render_main_page():
    st.title("💰 我的個人記帳本")
    st.caption("Streamlit + Pandas + Google Sheets + yfinance")

    spreadsheet = get_spreadsheet()
    tx_ws = ensure_worksheet(spreadsheet, WORKSHEET_NAME, tuple(TRANSACTIONS_HEADER))

    try:
        records = tx_ws.get_all_records()
        df = pd.DataFrame(records)
    except Exception as e:
        st.warning(f"目前尚未能讀取 Google Sheets 資料：{e}")
        df = pd.DataFrame()

    if not df.empty:
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
        df["Amount_TWD"] = pd.to_numeric(df.get("Amount_TWD"), errors="coerce")
        df["Date_parsed"] = pd.to_datetime(df["Date"], errors="coerce")
        df["YearMonth"] = df["Date_parsed"].dt.strftime("%Y-%m")
        for c in ("Item", "Note"):
            if c not in df.columns:
                df[c] = ""

    tab_overview, tab_assets, tab_charts, tab_budget, tab_detail = st.tabs(
        ["🏠 總覽", "📈 帳戶資產", "📊 圖表分析", "💰 預算", "📋 明細管理"]
    )
    with tab_overview:
        render_overview_tab(df, tx_ws)
    with tab_assets:
        render_assets_tab(df)
    with tab_charts:
        render_charts_tab(df)
    with tab_budget:
        render_budget_tab(df)
    with tab_detail:
        render_detail_tab(df, tx_ws)


def main():
    render_sidebar_form()
    render_recurring_manager()
    render_main_page()


if __name__ == "__main__":
    main()
