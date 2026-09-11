import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

from datetime import datetime
from zoneinfo import ZoneInfo
import os
import sqlalchemy as sa
import sys
import urllib.parse
import sqlalchemy as sa
# =========================
# 0. Kết nối Data Warehouse (SQL Server)
# =========================
# DW_CFG = {
#     "user": "risk",
#     "password": "Reiss##@#22",
#     "server": "192.168.10.163",
#     "database": "DWH-CoSo",
#     "driver": "ODBC Driver 17 for SQL Server",
# }
# DW_DSN = "DW_RISK"
#
# def get_dw_engine() -> sa.Engine:
#     conn_str = f"mssql+pyodbc:///?dsn={DW_DSN}"
#     print("[DW] Connecting with DSN:", conn_str)
#     engine = sa.create_engine(conn_str, pool_pre_ping=True)
#     return engine
#
# import urllib.parse
# import sqlalchemy as sa

def get_dw_engine() -> sa.Engine:
    # ⚠️ Thay SERVER=... bằng đúng chuỗi server bạn dùng khi Test ODBC
    # ví dụ: SERVER=192.168.10.163\RISK  hoặc SERVER=192.168.10.163  hoặc SERVER=risk
    odbc_str = (
        "DRIVER=ODBC Driver 17 for SQL Server;"
        "SERVER=192.168.10.163;"        # <-- copy y chang trong ODBC (nếu có \RISK thì để nguyên)
        "DATABASE=DWH-CoSo;"
        "UID=risk;"
        "PWD=Reiss##@#22;"
        "TrustServerCertificate=yes;"
    )

    params = urllib.parse.quote_plus(odbc_str)
    conn_str = f"mssql+pyodbc:///?odbc_connect={params}"
    print("[DW] Connecting with:", conn_str)

    engine = sa.create_engine(conn_str, pool_pre_ping=True)
    return engine
def get_latest_trade_date(engine: sa.Engine) -> str:
    """
    Lấy ngày giao dịch mới nhất (YYYY-MM-DD) trong số 4 bảng DW.
    """
    tables = [
        ("[DWH-CoSo].[dbo].[vpr0109]", "date"),
        ("[DWH-CoSo].[dbo].[230007]", "date"),
        ("[DWH-CoSo].[dbo].[Data_Hop_Infos]", "TRADING_DATE"),
        ("[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]", "Date"),
    ]

    dates = []
    for tbl, col in tables:
        sql = f"""
        SELECT MAX(CAST([{col}] AS date)) AS trade_date
        FROM {tbl}
        """
        df = pd.read_sql(sql, con=engine)
        d = df.loc[0, "trade_date"]
        if pd.notna(d):
            dates.append(d)

    if not dates:
        raise ValueError("Không tìm được trade_date nào trong DW")

    latest = max(dates)
    return latest.strftime("%Y-%m-%d")


# =========================
# Helper cho file "Data gửi Nam..." (Liquidity Deal Report)
# =========================

def _normalize_special_room_df(room_df: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hoá DataFrame special_room_df đọc từ Excel:

    - Nếu header đã có cột 'code' thì dùng luôn.
    - Nếu header có 'Stock' thì đổi sang 'code'.
    - Nếu không có cả hai, dò dòng nào chứa 'Stock' hoặc 'code'
      để làm header (trường hợp Liquidity Deal Report gốc).
    """
    df = room_df.copy()
    cols = list(df.columns)
    cols_stripped = [str(c).strip() for c in cols]

    # CASE 1: đã có cột 'code' trong header
    if "code" in cols_stripped:
        rename_map: dict = {}
        for c in cols:
            if str(c).strip() == "code":
                rename_map[c] = "code"
        if rename_map:
            df = df.rename(columns=rename_map)
        df["code"] = df["code"].astype(str).str.strip()
        return df

    # CASE 2: header đang có 'Stock' ngay từ đầu
    has_stock_col = any(c in ("Stock", "Stock ") for c in cols_stripped)
    if has_stock_col:
        rename_map: dict = {}
        for c in df.columns:
            if str(c).strip() in ("Stock", "Stock "):
                rename_map[c] = "code"
        if rename_map:
            df = df.rename(columns=rename_map)
        df["code"] = df["code"].astype(str).str.strip()
        return df

    # CASE 3: header nằm ở 1 dòng bên dưới (Liquidity Deal Report gốc)
    header_row_idx = None
    for idx, row in df.iterrows():
        if any(str(v).strip() in ("Stock", "code") for v in row.values):
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise KeyError(
            f"Không tìm thấy dòng header chứa 'Stock' hoặc 'code' trong special_room_df. "
            f"Columns hiện tại: {list(df.columns)}"
        )

    # Lấy dòng đó làm header mới
    header = df.iloc[header_row_idx]
    df = df.iloc[header_row_idx + 1 :].copy()
    df.columns = header

    # Gọi lại chính hàm này lần nữa để chạy qua CASE 1/CASE 2
    return _normalize_special_room_df(df)


def load_special_room_excel(path: str) -> pd.DataFrame:

    """
    Đọc file Excel special room (Data gửi Nam...) và chuẩn hoá
    về DataFrame có cột 'code', 'Used room'/ 'Used quantity today',
    'Maximum loan price', 'MR Approved Ratio (%)'.
    """
    raw_df = pd.read_excel(path)
    norm_df = _normalize_special_room_df(raw_df)
    return norm_df


# =========================
# 1. Core tính toán từ 4 DataFrame
# =========================

def _build_margin_report_from_frames(
    margin_df: pd.DataFrame,
    used_df: pd.DataFrame,
    vol_df: pd.DataFrame,
    price_df: pd.DataFrame,
    equity_vnd_million: float = 2_077_000.0,
    special_room_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Core logic tính toán report, dùng chung cho:
      - build_margin_report_df (Excel)
      - build_margin_report_df_from_dw (DW)
    """

    # 2. Chuẩn hoá tên cột để merge
    margin_df = margin_df.rename(columns={"Mã CK": "code"})
    used_df = used_df.rename(columns={"Mã chứng khoán": "code"})
    vol_df = vol_df.rename(columns={"Stock": "code"})
    price_df = price_df.rename(columns={"Mã": "code"})

    # 3. Lấy các cột cần thiết
    used_df = used_df[["code", "Room hệ thống đã sử dụng", "Room đặc biệt đã sử dụng"]]
    margin_df = margin_df[
        ["code", "Tỷ lệ vay KQ (%)", "Tỷ lệ vay TC  (%)", "Giá vay/Giá TSĐB tối đa (VND)"]
    ]
    vol_df = vol_df[["code", "Vol-listed"]]
    price_df = price_df[["code", "Đóng\n cửa"]]

    # 4. Merge lại
    df = (
        margin_df
        .merge(used_df, on="code", how="left")
        .merge(vol_df, on="code", how="left")
        .merge(price_df, on="code", how="left")
    )

    df["code"] = df["code"].astype(str).str.strip()
    df.reset_index(drop=True, inplace=True)

    # 5. Tính toán các chỉ tiêu cơ bản
    df["Số lượng chứng khoán niêm yết (3)"] = df["Vol-listed"] / 1000.0

    used_gr = df["Room hệ thống đã sử dụng"].fillna(0)
    used_sr = df["Room đặc biệt đã sử dụng"].fillna(0)
    total_used_shares = used_gr + used_sr

    limit_5pct_shares = 0.05 * df["Vol-listed"].fillna(0)
    allowed_lend_shares = np.minimum(
        total_used_shares / 1000.0,
        limit_5pct_shares / 1000.0,
    )

    df["Số lượng chứng khoán cho vay (2)"] = allowed_lend_shares
    df["Reference Price"] = df["Đóng\n cửa"] * 1000.0

    df["MR Loan ratio"] = df["Tỷ lệ vay KQ (%)"]
    df["DP Loan ratio"] = df["Tỷ lệ vay TC  (%)"]
    df["Max Price"] = df["Giá vay/Giá TSĐB tối đa (VND)"]
    df["Used GR"] = used_gr
    df["Used SR"] = used_sr
    df["Vốn chủ sở hữu của CTCK (4)"] = equity_vnd_million

    price_cap = df[["Reference Price", "Max Price"]].min(axis=1)
    mr_loan = df["MR Loan ratio"].fillna(0)
    dp_loan = df["DP Loan ratio"].fillna(0)
    ratio_cap = pd.concat([mr_loan, dp_loan], axis=1).max(axis=1)
    qty_room = used_gr + used_sr

    base_vnd = qty_room * price_cap * (ratio_cap / 100.0)

    vol_3 = df["Số lượng chứng khoán niêm yết (3)"].fillna(0)
    qty_2 = df["Số lượng chứng khoán cho vay (2)"].fillna(0)
    five_pct_qty = 0.05 * vol_3
    hit_5pct_mask = np.isclose(qty_2, five_pct_qty, rtol=1e-6, atol=1e-9)

    if hit_5pct_mask.any():
        base_vnd_cap = qty_room * price_cap * (ratio_cap / 100.0)
        base_vnd = base_vnd.where(~hit_5pct_mask, base_vnd_cap)

    base_debt_million = base_vnd.clip(lower=0) / 1_000_000.0
    max_debt_by_equity = 0.10 * df["Vốn chủ sở hữu của CTCK (4)"]
    base_debt_million = np.minimum(base_debt_million, max_debt_by_equity)

    # Khởi tạo các cột special
    df["Total Used room"] = np.nan
    df["Giá chặn riêng"] = np.nan
    df["Tỷ lệ đặc biệt"] = np.nan  # MR Approved Ratio (%)

    # ===== GHÉP ROOM ĐẶC BIỆT (special_room_df) =====
    if special_room_df is not None:
        # Chuẩn hoá header và cột 'code'
        room_df = _normalize_special_room_df(special_room_df)

        cols = list(room_df.columns)

        # Xác định cột used, maximum loan price, approved ratio
        used_col = None
        for candidate in ["Used quantity today", "Used room", "Total Used room"]:
            if candidate in cols:
                used_col = candidate
                break


        max_price_col = None
        for c in cols:
            if "Maximum loan price" in str(c):
                max_price_col = c
                break

        ratio_col = None
        for c in cols:
            if "MR Approved Ratio" in str(c):
                ratio_col = c
                break

        # Chuẩn numeric
        for col in [used_col, max_price_col, ratio_col]:
            if col is not None and col in room_df.columns:
                room_df[col] = pd.to_numeric(room_df[col], errors="coerce").fillna(0)

        # Gom theo code
        agg_kwargs: dict = {}
        if used_col:
            agg_kwargs["Special_used_qty"] = (used_col, "sum")
        if max_price_col:
            agg_kwargs["Gia_chan_rieng"] = (max_price_col, "min")
        if ratio_col:
            agg_kwargs["Special_rate"] = (ratio_col, "max")

        if not agg_kwargs:
            # Có code nhưng lại không có cột numeric nào dùng được
            raise KeyError(
                f"Không tìm được cột Used room / Maximum loan price / MR Approved Ratio trong special_room_df. "
                f"Columns: {cols}"
            )

        agg_room = room_df.groupby("code", as_index=False).agg(**agg_kwargs)

        # Merge vào df
        df = df.merge(agg_room, on="code", how="left")

        if "Special_used_qty" in df.columns:
            df["Total Used room"] = df["Special_used_qty"]
        if "Gia_chan_rieng" in df.columns:
            df["Giá chặn riêng"] = df["Gia_chan_rieng"]
        if "Special_rate" in df.columns:
            df["Tỷ lệ đặc biệt"] = df["Special_rate"].fillna(0)

    # ===== DƯ NỢ (1) CHUNG + CASE ĐẶC BIỆT =====
    df["Dư nợ cho vay GDKQ (1)"] = base_debt_million

    special_qty = df["Total Used room"].fillna(0)
    gia_chan = df["Giá chặn riêng"].fillna(df["Max Price"])
    special_rate = df["Tỷ lệ đặc biệt"].fillna(0)
    ref_price = df["Reference Price"].fillna(0)
    max_price = df["Max Price"].fillna(0)

    base_price_cap = price_cap
    special_price_cap = pd.concat([gia_chan, ref_price], axis=1).min(axis=1)

    is_special = special_qty > 0
    cond_price_diff = (gia_chan != ref_price) & (gia_chan != max_price)
    cond_price_eq_max = gia_chan == max_price

    million_factor = 1_000_000.0

    # Case 1
    case1_mask = is_special & (special_rate == mr_loan) & cond_price_diff
    if case1_mask.any():
        part1_vnd = (used_sr - special_qty).clip(lower=0) * base_price_cap * ratio_cap / 100.0
        part2_vnd = used_gr * base_price_cap * ratio_cap / 100.0
        part3_vnd = special_qty * special_price_cap * ratio_cap / 100.0
        special_vnd = part1_vnd + part2_vnd + part3_vnd
        special_debt_million = special_vnd.clip(lower=0) / million_factor
        special_debt_million = np.minimum(special_debt_million, max_debt_by_equity)
        df.loc[case1_mask, "Dư nợ cho vay GDKQ (1)"] = special_debt_million[case1_mask]

    # Case 2
    case2_mask = (
        is_special
        & (special_rate != mr_loan)
        & cond_price_diff
        & (used_sr > special_qty)
    )
    if case2_mask.any():
        part1_vnd = used_gr * base_price_cap * ratio_cap / 100.0
        part2_vnd = (used_sr - special_qty).clip(lower=0) * base_price_cap * ratio_cap / 100.0
        part3_vnd = special_qty * special_price_cap * special_rate / 100.0
        special_vnd = part1_vnd + part2_vnd + part3_vnd
        special_debt_million = special_vnd.clip(lower=0) / million_factor
        special_debt_million = np.minimum(special_debt_million, max_debt_by_equity)
        df.loc[case2_mask, "Dư nợ cho vay GDKQ (1)"] = special_debt_million[case2_mask]

    # Case 3
    case3_mask = (
        is_special
        & (special_rate != mr_loan)
        & cond_price_diff
        & (used_sr == special_qty)
        & (special_qty > 0)
    )
    if case3_mask.any():
        part1_vnd = used_gr * base_price_cap * ratio_cap / 100.0
        part2_vnd = used_sr * special_price_cap * special_rate / 100.0
        special_vnd = part1_vnd + part2_vnd
        special_debt_million = special_vnd.clip(lower=0) / million_factor
        special_debt_million = np.minimum(special_debt_million, max_debt_by_equity)
        df.loc[case3_mask, "Dư nợ cho vay GDKQ (1)"] = special_debt_million[case3_mask]

    # Case 4
    case4_mask = is_special & (special_rate != mr_loan) & cond_price_eq_max
    if case4_mask.any():
        part1_vnd = used_gr * base_price_cap * ratio_cap / 100.0
        part2_vnd = used_sr * special_price_cap * special_rate / 100.0
        special_vnd = part1_vnd + part2_vnd
        special_debt_million = special_vnd.clip(lower=0) / million_factor
        special_debt_million = np.minimum(special_debt_million, max_debt_by_equity)
        df.loc[case4_mask, "Dư nợ cho vay GDKQ (1)"] = special_debt_million[case4_mask]

    already_special_mask = case1_mask | case2_mask | case3_mask | case4_mask
    sr_only_mask = (
        (used_gr.fillna(0) == 0)
        & (used_sr.fillna(0) > 0)
        & df["Total Used room"].notna()
        & (~already_special_mask)
    )

    if sr_only_mask.any():
        total_used_room = df["Total Used room"].fillna(0)
        price_cap_sr = df[["Max Price", "Reference Price"]].min(axis=1)
        ratio_cap_sr = df[["MR Loan ratio", "DP Loan ratio"]].max(axis=1)
        sr_only_vnd = total_used_room * price_cap_sr * ratio_cap_sr / 100.0
        sr_only_debt_million = sr_only_vnd.clip(lower=0) / million_factor
        sr_only_debt_million = np.minimum(sr_only_debt_million, max_debt_by_equity)
        df.loc[sr_only_mask, "Dư nợ cho vay GDKQ (1)"] = sr_only_debt_million[sr_only_mask]

    df["Tỷ lệ dư nợ/VCSH (1)/(4)"] = (
        df["Dư nợ cho vay GDKQ (1)"] / df["Vốn chủ sở hữu của CTCK (4)"]
    )
    df["Tỷ lệ CK cho vay/CKNY (2)/(3)"] = (
        df["Số lượng chứng khoán cho vay (2)"] / df["Số lượng chứng khoán niêm yết (3)"]
    )

    ordered_cols = [
        "code",
        "Dư nợ cho vay GDKQ (1)",
        "Số lượng chứng khoán cho vay (2)",
        "Số lượng chứng khoán niêm yết (3)",
        "Vốn chủ sở hữu của CTCK (4)",
        "Tỷ lệ dư nợ/VCSH (1)/(4)",
        "Tỷ lệ CK cho vay/CKNY (2)/(3)",
        "Used GR",
        "Used SR",
        "Reference Price",
        "MR Loan ratio",
        "DP Loan ratio",
        "Max Price",
        "Total Used room",
        "Giá chặn riêng",
        "Tỷ lệ đặc biệt",
    ]

    df_out = df[ordered_cols].copy()
    df_out = df_out.sort_values("code").reset_index(drop=True)
    df_out.insert(0, "STT", range(1, len(df_out) + 1))

    return df_out


# =========================
# 1A. Đọc & tính toán số liệu từ EXCEL
# =========================

def build_margin_report_df(
    excel_path: str,
    equity_vnd_million: float = 2_077_000.0,
    special_room_path: str | None = None,
) -> pd.DataFrame:
    """
    Đọc file raw Excel (Margin List, 230007, Vol list, Matched result)
    và trả về DataFrame báo cáo.
    """

    margin_df = pd.read_excel(excel_path, sheet_name="Margin List")
    used_df = pd.read_excel(excel_path, sheet_name="230007")
    vol_df = pd.read_excel(excel_path, sheet_name="Vol list")
    price_df = pd.read_excel(excel_path, sheet_name="Matched result")

    special_room_df = None
    if special_room_path is not None:
        special_room_df = load_special_room_excel(special_room_path)

    df_out = _build_margin_report_from_frames(
        margin_df=margin_df,
        used_df=used_df,
        vol_df=vol_df,
        price_df=price_df,
        equity_vnd_million=equity_vnd_million,
        special_room_df=special_room_df,
    )
    return df_out


# =========================
# 1B. Đọc & tính toán số liệu từ DATA WAREHOUSE (DW)
# =========================

def build_margin_report_df_from_dw(
    engine: sa.Engine,
    trade_date: str,
    equity_vnd_million: float = 2_077_000.0,
    margin_table: str = "[DWH-CoSo].[dbo].[vpr0109]",
    used_table: str = "[DWH-CoSo].[dbo].[230007]",
    vol_table: str = "[DWH-CoSo].[dbo].[Data_Hop_Infos]",
    price_table: str = "[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]",
    special_room_table = "[DWH-CoSo].[dbo].[vpr0108]",
    special_room_excel_path: str | None = None,
) -> tuple[pd.DataFrame, set[str]]:
    """
    Đọc trực tiếp DW, có thể kết hợp special room từ Excel hoặc bảng DW.
    trade_date: 'YYYY-MM-DD'
    """

    # ----- 1. Margin list -----
    margin_sql = f"""
    SELECT
        m.ticker_code          AS [Mã CK],
        m.collateral_ratio     AS [Tỷ lệ vay KQ (%)],
        m.margin_ratio         AS [Tỷ lệ vay TC  (%)],
        m.collateral_max_price AS [Giá vay/Giá TSĐB tối đa (VND)]
    FROM {margin_table} AS m
    WHERE CAST(m.[date] AS date) = '{trade_date}'
    AND m.room_code = 'CL01_PHS'
    """
    margin_df = pd.read_sql(margin_sql, con=engine)

    # ----- 2. Used room -----
    used_sql = f"""
    SELECT
        u.ticker            AS [Mã chứng khoán],
        u.system_used_room      AS [Room hệ thống đã sử dụng],
        u.used_special_room AS [Room đặc biệt đã sử dụng]
    FROM {used_table} AS u
    WHERE CAST(u.[date] AS date) = '{trade_date}'
    """
    used_df = pd.read_sql(used_sql, con=engine)

    # ----- 3. Vol listed -----
    vol_sql = f"""
    SELECT
        v.SYMBOL     AS [Stock],
        v.VOL_LISTED AS [Vol-listed]
    FROM {vol_table} AS v
    WHERE CAST(v.[TRADING_DATE] AS date) = '{trade_date}'
    """
    vol_df = pd.read_sql(vol_sql, con=engine)

    # ----- 4. Giá đóng cửa -----
    price_sql = f"""
    SELECT
        p.Ticker AS [Mã],
        p.[Close]  AS [Đóng
 cửa]
    FROM {price_table} AS p
    WHERE CAST(p.[Date] AS date) = '{trade_date}'
    """
    price_df = pd.read_sql(price_sql, con=engine)

    # ----- 5. Special room (Excel hoặc bảng DW) -----
    special_room_df = None

    # 5A. Nếu có file Excel room đặc biệt thì ưu tiên dùng
    if special_room_excel_path:
        if os.path.exists(special_room_excel_path):
            print(f"[INFO] Đọc special room từ Excel: {special_room_excel_path}")
            special_room_df = load_special_room_excel(special_room_excel_path)
        else:
            print(f"[WARN] Không tìm thấy file special room: {special_room_excel_path}")
            special_room_df = None

    # 5B. Nếu không có Excel mà có bảng DW thì dùng bảng
        # 5B. Nếu không có Excel mà có bảng DW thì dùng join vpr0109 + vpr0108
    elif special_room_table:
        special_sql = f"""
        WITH special_map AS (
            SELECT *
            FROM (VALUES
                ('TV4', 'CL08_PHS', '0095'),
                ('GEX', 'CL07_PHS', '0074'),
                ('BAF', 'CL47_PHS', '0216'),
                ('DL1', 'CL59_PHS', '0287'),
                ('SAM', 'CL16_PHS', '0084'),
                ('C69', 'CL44_PHS', '0166'),
                ('HVH', 'CL51_PHS', '0272'),
                ('HHP', 'CL58_PHS', '0215'),
                ('SJS', 'CL57_PHS', '0222'),
                ('SJS', 'CL57_PHS', '0282'),
                ('SJS', 'CL57_PHS', '0283'),
                ('CKG', 'CL20_PHS', '0219'),
                ('SBT', 'CL52_PHS', '0275'),
                ('CDC', 'CL56_PHS', '0280')
            ) AS v(ticker, room_code_09, room_code_08)
        )
        SELECT
            m.ticker_code      AS [Stock],
            u.used_volume      AS [Used room],
            m.margin_max_price AS [Maximum loan price],
            m.margin_ratio     AS [MR Approved Ratio (%)]
        FROM special_map s
        JOIN {margin_table} AS m
          ON m.ticker_code = s.ticker
         AND m.room_code   = s.room_code_09
         AND CAST(m.[date] AS date) = '{trade_date}'
        JOIN {special_room_table} AS u
          ON u.ticker    = s.ticker
         AND u.room_code = s.room_code_08
         AND CAST(u.[date] AS date) = '{trade_date}'
        """

        special_room_df = pd.read_sql(special_sql, con=engine)


    # ----- 6. Build full report frame -----
    df_full = _build_margin_report_from_frames(
        margin_df=margin_df,
        used_df=used_df,
        vol_df=vol_df,
        price_df=price_df,
        equity_vnd_million=equity_vnd_million,
        special_room_df=special_room_df,
    )

    # ----- 7. Lấy danh sách mã có room đặc biệt để tô màu -----
    special_codes = set(
        df_full.loc[df_full["Total Used room"].fillna(0) > 0, "code"]
        .astype(str)
        .str.strip()
    )

    return df_full, special_codes

# =========================
# 2. Helper: num_or_dash
# =========================

def num_or_dash(x):
    """Trả về '-' nếu x = 0 hoặc NaN, ngược lại trả về float(x)."""
    if pd.isna(x):
        return "-"
    try:
        v = float(x)
    except (TypeError, ValueError):
        return x
    return "-" if abs(v) < 1e-9 else v


# =========================
# 2A. File 1 từ EXCEL: Copy raw + thêm sheet Dư nợ theo used room_PY
# =========================

def write_margin_report_sheet(
    excel_input_path: str,
    excel_output_path: str,
    sheet_name: str = "Dư nợ theo used room_PY",
    equity_vnd_million: float = 2_077_000.0,
    special_room_path: str | None = None,
):
    """
    BẢN CŨ (Excel):
    - Đọc file input (raw)
    - Tính DataFrame báo cáo margin
    - Tạo workbook mới với 1 sheet report (giống kiểu 'Dư nợ theo used room')
    - Lưu ra excel_output_path
    """

    # 1. Tính toán DataFrame
    df = build_margin_report_df(
        excel_input_path,
        equity_vnd_million=equity_vnd_million,
        special_room_path=special_room_path,
    )

    # >>> LẤY DANH SÁCH MÃ ĐẶC BIỆT TỪ FILE 1 SỐ ROOM ĐẶC BIỆT
    special_codes: set[str] = set()
    if special_room_path is not None:
        room_df = load_special_room_excel(special_room_path)
        special_codes = {
            str(x).strip()
            for x in room_df["code"].dropna().astype(str)
        }

    # 2. Tạo workbook mới
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Nếu sheet này đã tồn tại thì xoá trước (tránh trùng)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    # 3. Header giống logic 'Dư nợ theo used room'
    headers = [
        "STT",
        "Mã chứng khoán",
        "Dư nợ cho vay GDKQ \n(1)",
        "Số lượng chứng khoán cho vay của CTCK \n(2)",
        "Số lượng chứng khoán niêm yết của TCNY\n(3)",
        "Vốn chủ sở hữu của CTCK \n(4)",
        "Tỷ lệ dư nợ/VCSH \n(1)/(4)",
        "Tỷ lệ CK cho vay/CKNY\n(2)/(3)",
        "Used GR",
        "Used SR",
        "Reference Price",
        "MR Loan ratio",
        "DP Loan ratio",
        "Max Price",
        "Control",
        "Total Used room",
        "Giá chặn riêng",
        "Tỷ lệ đặc biệt",
    ]
    n_cols = len(headers)

    # 4. Tiêu đề + dòng đơn vị
    ws.cell(row=1, column=1, value="TÌNH HÌNH GIAO DỊCH KÝ QUỸ")
    ws.cell(row=2, column=1, value="Đơn vị : nghìn cổ phiếu,triệu đồng")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    # 5. Header ở hàng 3
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col_idx, value=header)

    # 6. Ghi dữ liệu (hàng 4 trở đi)
    for row_idx, row in df.iterrows():
        excel_row = row_idx + 4

        ws.cell(row=excel_row, column=1, value=int(row["STT"]))
        ws.cell(row=excel_row, column=2, value=row["code"])

        ws.cell(row=excel_row, column=3, value=num_or_dash(row["Dư nợ cho vay GDKQ (1)"]))
        ws.cell(row=excel_row, column=4, value=num_or_dash(row["Số lượng chứng khoán cho vay (2)"]))
        ws.cell(row=excel_row, column=5, value=num_or_dash(row["Số lượng chứng khoán niêm yết (3)"]))

        ws.cell(row=excel_row, column=6, value=num_or_dash(row["Vốn chủ sở hữu của CTCK (4)"]))

        ws.cell(
            row=excel_row,
            column=7,
            value=float(row["Tỷ lệ dư nợ/VCSH (1)/(4)"]) if pd.notna(row["Tỷ lệ dư nợ/VCSH (1)/(4)"]) else None,
        )
        ws.cell(
            row=excel_row,
            column=8,
            value=float(row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]) if pd.notna(
                row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]) else None,
        )

        ws.cell(row=excel_row, column=9, value=num_or_dash(row["Used GR"]))
        ws.cell(row=excel_row, column=10, value=num_or_dash(row["Used SR"]))
        ws.cell(row=excel_row, column=11, value=num_or_dash(row["Reference Price"]))
        ws.cell(row=excel_row, column=12, value=num_or_dash(row["MR Loan ratio"]))
        ws.cell(row=excel_row, column=13, value=num_or_dash(row["DP Loan ratio"]))
        ws.cell(row=excel_row, column=14, value=num_or_dash(row["Max Price"]))

        ws.cell(
            row=excel_row,
            column=16,
            value=num_or_dash(row.get("Total Used room", np.nan)),
        )
        ws.cell(
            row=excel_row,
            column=17,
            value=num_or_dash(row.get("Giá chặn riêng", np.nan)),
        )
        ws.cell(
            row=excel_row,
            column=18,
            value=row.get("Tỷ lệ đặc biệt", 0),
        )

    # 6b. Hàng tổng
    total_row = 4 + len(df)
    total_debt = df["Dư nợ cho vay GDKQ (1)"].sum()
    total_qty = df["Số lượng chứng khoán cho vay (2)"].sum()

    ws.cell(row=total_row, column=3, value=num_or_dash(total_debt))
    ws.cell(row=total_row, column=4, value=num_or_dash(total_qty))

    # =========================
    # 7. Format
    # =========================

    title_font = Font(bold=True, size=13)
    unit_font = Font(italic=True)
    header_font = Font(bold=True, color="000000")
    header_fill = PatternFill(start_color="FFC867", end_color="FFC867", fill_type="solid")
    value_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    special_fill = PatternFill(start_color="FFFF66", end_color="FFFF66", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].font = unit_font
    ws["A2"].alignment = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.alignment = center

    for col in range(9, 15):
        ws.cell(row=3, column=col).fill = header_fill

    ws.cell(row=3, column=15).fill = special_fill

    first_header_row = 3
    first_data_row = 4
    last_data_row = 3 + len(df)

    for row in ws.iter_rows(
        min_row=first_header_row,
        max_row=first_header_row,
        min_col=1,
        max_col=14,
    ):
        for cell in row:
            cell.border = border

    for row in ws.iter_rows(
        min_row=first_data_row,
        max_row=last_data_row,
        min_col=1,
        max_col=14,
    ):
        for cell in row:
            cell.border = border

    bold_font = Font(bold=True)
    for col_idx in (3, 4):
        ws.cell(row=total_row, column=col_idx).font = bold_font

    for row in ws.iter_rows(
        min_row=first_data_row,
        max_row=last_data_row,
        min_col=9,
        max_col=14,
    ):
        for cell in row:
            cell.fill = value_fill

    special_mask = df["Total Used room"].fillna(0) > 0
    for row_idx, is_sp in enumerate(special_mask, start=0):
        if not is_sp:
            continue
        excel_row = first_data_row + row_idx
        for col_idx in range(2, 19):
            ws.cell(row=excel_row, column=col_idx).fill = special_fill

    width_map = {
        1: 5, 2: 12, 3: 12, 4: 12,
        5: 15, 6: 12, 7: 12, 8: 12,
        9: 12, 10: 12, 11: 12, 12: 12,
        13: 12, 14: 12, 15: 12, 16: 12, 17: 12, 18: 12,
    }
    for col_idx, width in width_map.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=1, max_col=n_cols):
        row[2].number_format = "#,##0"
        row[3].number_format = "#,##0"
        row[4].number_format = "#,##0"
        row[5].number_format = "#,##0"
        row[6].number_format = "0.00%"
        row[7].number_format = "0.00%"
        row[8].number_format = "#,##0"
        row[9].number_format = "#,##0"
        row[10].number_format = "#,##0"
        row[13].number_format = "#,##0"
        row[11].number_format = "0"
        row[12].number_format = "0"
        row[15].number_format = "#,##0"
        row[16].number_format = "#,##0"
        row[17].number_format = "0"  # Tỷ lệ đặc biệt (col 18)

    right_align = Alignment(horizontal="right", vertical="center")
    for row in ws.iter_rows(
        min_row=4,
        max_row=total_row,
        min_col=3,
        max_col=18,
    ):
        for cell in row:
            cell.alignment = right_align

    ws.freeze_panes = "I4"

    if special_codes:
        special_fill2 = PatternFill(
            start_color="FFFF66",
            end_color="FFFF66",
            fill_type="solid",
        )
        for row_idx, row in df.iterrows():
            code = str(row["code"]).strip()
            if code in special_codes:
                excel_row = row_idx + 4
                for col_idx in range(2, 19):
                    ws.cell(row=excel_row, column=col_idx).fill = special_fill2

    max_row = total_row
    max_col = n_cols
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                new_font = copy(cell.font)
                new_font.name = "Times New Roman"
                cell.font = new_font

    wb.save(excel_output_path)
    print(f"Da tao file moi '{excel_output_path}' voi sheet '{sheet_name}'")


# =========================
# 2A'. BẢN FROM_DF: dùng cho DW (df_full đã có sẵn)
# =========================

def write_margin_report_sheet_from_df(
    df: pd.DataFrame,
    special_codes: set[str] | None,
    excel_output_path: str,
    sheet_name: str = "Dư nợ theo used room",
):
    """
    Phiên bản dùng khi đã có df_full (từ DW), không cần đọc Excel input nữa.
    Logic format giống write_margin_report_sheet.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    headers = [
        "STT",
        "Mã chứng khoán",
        "Dư nợ cho vay GDKQ \n(1)",
        "Số lượng chứng khoán cho vay của CTCK \n(2)",
        "Số lượng chứng khoán niêm yết của TCNY\n(3)",
        "Vốn chủ sở hữu của CTCK \n(4)",
        "Tỷ lệ dư nợ/VCSH \n(1)/(4)",
        "Tỷ lệ CK cho vay/CKNY\n(2)/(3)",
        "Used GR",
        "Used SR",
        "Reference Price",
        "MR Loan ratio",
        "DP Loan ratio",
        "Max Price",
        "Control",
        "Total Used room",
        "Giá chặn riêng",
        "Tỷ lệ đặc biệt",
    ]
    n_cols = len(headers)

    ws.cell(row=1, column=1, value="TÌNH HÌNH GIAO DỊCH KÝ QUỸ")
    ws.cell(row=2, column=1, value="Đơn vị : nghìn cổ phiếu,triệu đồng")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col_idx, value=header)

    for row_idx, row in df.iterrows():
        excel_row = row_idx + 4

        ws.cell(row=excel_row, column=1, value=int(row["STT"]))
        ws.cell(row=excel_row, column=2, value=row["code"])

        ws.cell(row=excel_row, column=3, value=num_or_dash(row["Dư nợ cho vay GDKQ (1)"]))
        ws.cell(row=excel_row, column=4, value=num_or_dash(row["Số lượng chứng khoán cho vay (2)"]))
        ws.cell(row=excel_row, column=5, value=num_or_dash(row["Số lượng chứng khoán niêm yết (3)"]))
        ws.cell(row=excel_row, column=6, value=num_or_dash(row["Vốn chủ sở hữu của CTCK (4)"]))

        ws.cell(
            row=excel_row,
            column=7,
            value=float(row["Tỷ lệ dư nợ/VCSH (1)/(4)"]) if pd.notna(row["Tỷ lệ dư nợ/VCSH (1)/(4)"]) else None,
        )
        ws.cell(
            row=excel_row,
            column=8,
            value=float(row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]) if pd.notna(
                row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]) else None,
        )

        ws.cell(row=excel_row, column=9, value=num_or_dash(row["Used GR"]))
        ws.cell(row=excel_row, column=10, value=num_or_dash(row["Used SR"]))
        ws.cell(row=excel_row, column=11, value=num_or_dash(row["Reference Price"]))
        ws.cell(row=excel_row, column=12, value=num_or_dash(row["MR Loan ratio"]))
        ws.cell(row=excel_row, column=13, value=num_or_dash(row["DP Loan ratio"]))
        ws.cell(row=excel_row, column=14, value=num_or_dash(row["Max Price"]))

        ws.cell(row=excel_row, column=16, value=num_or_dash(row.get("Total Used room", np.nan)))
        ws.cell(row=excel_row, column=17, value=num_or_dash(row.get("Giá chặn riêng", np.nan)))
        ws.cell(row=excel_row, column=18, value=row.get("Tỷ lệ đặc biệt", 0))

    total_row = 4 + len(df)
    total_debt = df["Dư nợ cho vay GDKQ (1)"].sum()
    total_qty = df["Số lượng chứng khoán cho vay (2)"].sum()

    ws.cell(row=total_row, column=3, value=num_or_dash(total_debt))
    ws.cell(row=total_row, column=4, value=num_or_dash(total_qty))

    title_font = Font(bold=True, size=13)
    unit_font = Font(italic=True)
    header_font = Font(bold=True, color="000000")
    header_fill = PatternFill(start_color="FFC867", end_color="FFC867", fill_type="solid")
    value_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    special_fill = PatternFill(start_color="FFFF66", end_color="FFFF66", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].font = unit_font
    ws["A2"].alignment = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.alignment = center

    for col in range(9, 15):
        ws.cell(row=3, column=col).fill = header_fill
    ws.cell(row=3, column=15).fill = special_fill

    first_header_row = 3
    first_data_row = 4
    last_data_row = 3 + len(df)

    for row in ws.iter_rows(min_row=first_header_row, max_row=first_header_row, min_col=1, max_col=14):
        for cell in row:
            cell.border = border

    for row in ws.iter_rows(min_row=first_data_row, max_row=last_data_row, min_col=1, max_col=14):
        for cell in row:
            cell.border = border

    bold_font = Font(bold=True)
    for col_idx in (3, 4):
        ws.cell(row=total_row, column=col_idx).font = bold_font

    for row in ws.iter_rows(min_row=first_data_row, max_row=last_data_row, min_col=9, max_col=14):
        for cell in row:
            cell.fill = value_fill

    special_mask = df["Total Used room"].fillna(0) > 0
    for row_idx, is_sp in enumerate(special_mask, start=0):
        if not is_sp:
            continue
        excel_row = first_data_row + row_idx
        for col_idx in range(2, 19):
            ws.cell(row=excel_row, column=col_idx).fill = special_fill

    width_map = {
        1: 5, 2: 12, 3: 12, 4: 12,
        5: 15, 6: 12, 7: 12, 8: 12,
        9: 12, 10: 12, 11: 12, 12: 12,
        13: 12, 14: 12, 15: 12, 16: 12, 17: 12, 18: 12,
    }
    for col_idx, width in width_map.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=1, max_col=n_cols):
        row[2].number_format = "#,##0"
        row[3].number_format = "#,##0"
        row[4].number_format = "#,##0"
        row[5].number_format = "#,##0"
        row[6].number_format = "0.00%"
        row[7].number_format = "0.00%"
        row[8].number_format = "#,##0"
        row[9].number_format = "#,##0"
        row[10].number_format = "#,##0"
        row[13].number_format = "#,##0"
        row[11].number_format = "0"
        row[12].number_format = "0"
        row[15].number_format = "#,##0"
        row[16].number_format = "#,##0"
        row[17].number_format = "0"

    right_align = Alignment(horizontal="right", vertical="center")
    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=3, max_col=18):
        for cell in row:
            cell.alignment = right_align

    ws.freeze_panes = "I4"

    if special_codes:
        special_fill2 = PatternFill(start_color="FFFF66", end_color="FFFF66", fill_type="solid")
        for row_idx, row in df.iterrows():
            code = str(row["code"]).strip()
            if code in special_codes:
                excel_row = row_idx + 4
                for col_idx in range(2, 19):
                    ws.cell(row=excel_row, column=col_idx).fill = special_fill2

    max_row = total_row
    max_col = n_cols
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                new_font = copy(cell.font)
                new_font.name = "Times New Roman"
                cell.font = new_font

    wb.save(excel_output_path)
    print(f"Da tao file tu DF: '{excel_output_path}' voi sheet '{sheet_name}'")


# =========================
# 2B. File 2 từ EXCEL: Tạo file '_PY.xlsx' báo cáo ngày
# =========================

def write_margin_report_file(
    raw_excel_path: str,
    output_report_path: str,
    sheet_name: str = "1.f_thgdkq_06692",
    equity_vnd_million: float = 2_077_000.0,
    special_room_path: str | None = None,
):
    """
    - Đọc file RAW
    - Tính DataFrame báo cáo margin
    - Tạo workbook mới với 1 sheet, format giống báo cáo ngày.
    """
    df_full = build_margin_report_df(
        raw_excel_path,
        equity_vnd_million=equity_vnd_million,
        special_room_path=special_room_path,
    )

    df = df_full[
        [
            "STT",
            "code",
            "Dư nợ cho vay GDKQ (1)",
            "Số lượng chứng khoán cho vay (2)",
            "Số lượng chứng khoán niêm yết (3)",
            "Vốn chủ sở hữu của CTCK (4)",
            "Tỷ lệ dư nợ/VCSH (1)/(4)",
            "Tỷ lệ CK cho vay/CKNY (2)/(3)",
        ]
    ].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = [
        "STT",
        "Mã chứng khoán",
        "Dư nợ cho vay GDKQ (1)",
        "Số lượng chứng khoán cho vay của CTCK (2)",
        "Số lượng chứng khoán niêm yết của TCNY (3)",
        "Vốn chủ sở hữu của CTCK (4)",
        "Tỷ lệ dư nợ/VCSH (1)/(4)",
        "Tỷ lệ CK cho vay/CKNY (2)/(3)",
    ]
    n_cols = len(headers)

    ws.cell(row=1, column=1, value="TÌNH HÌNH GIAO DỊCH KÝ QUỸ")
    ws.cell(row=2, column=1, value="Đơn vị: nghìn cổ phiếu, triệu đồng")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col_idx, value=header)

    first_data_row = 4
    for row_idx, row in df.iterrows():
        excel_row = first_data_row + row_idx

        ws.cell(row=excel_row, column=1, value=int(row["STT"]))
        ws.cell(row=excel_row, column=2, value=row["code"])

        ws.cell(row=excel_row, column=3, value=num_or_dash(row["Dư nợ cho vay GDKQ (1)"]))
        ws.cell(row=excel_row, column=4, value=num_or_dash(row["Số lượng chứng khoán cho vay (2)"]))
        ws.cell(row=excel_row, column=5, value=num_or_dash(row["Số lượng chứng khoán niêm yết (3)"]))
        ws.cell(row=excel_row, column=6, value=num_or_dash(row["Vốn chủ sở hữu của CTCK (4)"]))

        v1 = row["Tỷ lệ dư nợ/VCSH (1)/(4)"]
        v2 = row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]
        ws.cell(
            row=excel_row,
            column=7,
            value=float(v1) if pd.notna(v1) else None,
        )
        ws.cell(
            row=excel_row,
            column=8,
            value=float(v2) if pd.notna(v2) else None,
        )

    total_row = first_data_row + len(df)
    last_data_row = first_data_row + len(df) - 1

    total_debt = df["Dư nợ cho vay GDKQ (1)"].sum()
    total_qty = df["Số lượng chứng khoán cho vay (2)"].sum()

    ws.cell(row=total_row, column=2, value="Tổng cộng")
    ws.cell(row=total_row, column=3, value=num_or_dash(total_debt))
    ws.cell(row=total_row, column=4, value=num_or_dash(total_qty))

    title_font = Font(bold=True, size=16)
    unit_font = Font(italic=True, size=13)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].font = unit_font
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for row in ws.iter_rows(
        min_row=3,
        max_row=last_data_row,
        min_col=1,
        max_col=n_cols,
    ):
        for cell in row:
            cell.border = border

    for col_idx in range(2, 5):
        ws.cell(row=total_row, column=col_idx).border = border

    bold_font = Font(bold=True)
    for col_idx in (3, 4):
        ws.cell(row=total_row, column=col_idx).font = bold_font

    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=1, max_col=n_cols):
        row[2].number_format = "#,##0"
        row[3].number_format = "#,##0"
        row[4].number_format = "#,##0"
        row[5].number_format = "#,##0"
        row[6].number_format = "0.00%"
        row[7].number_format = "0.00%"

    for row in ws.iter_rows(
        min_row=4,
        max_row=total_row,
        min_col=3,
        max_col=8,
    ):
        for cell in row:
            cell.alignment = right_align

    width_map = {
        1: 5,
        2: 25,
        3: 25,
        4: 25,
        5: 25,
        6: 25,
        7: 25,
        8: 25,
    }
    for col_idx, width in width_map.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A4"

    max_row = total_row
    max_col = n_cols
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                new_font = copy(cell.font)
                new_font.name = "Times New Roman"
                cell.font = new_font

    wb.save(output_report_path)
    print(f"Da tao file bao cao ngay: {output_report_path}")


# =========================
# 2B'. BẢN FROM_DF: report ngày dùng cho DW
# =========================

def write_margin_report_file_from_df(
    df_full: pd.DataFrame,
    output_report_path: str,
    sheet_name: str = "1.f_thgdkq_06692",
):
    """
    Bản dùng khi đã có df_full từ DW (không đọc Excel).
    Format giống write_margin_report_file.
    """
    df = df_full[
        [
            "STT",
            "code",
            "Dư nợ cho vay GDKQ (1)",
            "Số lượng chứng khoán cho vay (2)",
            "Số lượng chứng khoán niêm yết (3)",
            "Vốn chủ sở hữu của CTCK (4)",
            "Tỷ lệ dư nợ/VCSH (1)/(4)",
            "Tỷ lệ CK cho vay/CKNY (2)/(3)",
        ]
    ].copy()

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = [
        "STT",
        "Mã chứng khoán",
        "Dư nợ cho vay GDKQ (1)",
        "Số lượng chứng khoán cho vay của CTCK (2)",
        "Số lượng chứng khoán niêm yết của TCNY (3)",
        "Vốn chủ sở hữu của CTCK (4)",
        "Tỷ lệ dư nợ/VCSH (1)/(4)",
        "Tỷ lệ CK cho vay/CKNY (2)/(3)",
    ]
    n_cols = len(headers)

    ws.cell(row=1, column=1, value="TÌNH HÌNH GIAO DỊCH KÝ QUỸ")
    ws.cell(row=2, column=1, value="Đơn vị: nghìn cổ phiếu, triệu đồng")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=3, column=col_idx, value=header)

    first_data_row = 4
    for row_idx, row in df.iterrows():
        excel_row = first_data_row + row_idx

        ws.cell(row=excel_row, column=1, value=int(row["STT"]))
        ws.cell(row=excel_row, column=2, value=row["code"])

        ws.cell(row=excel_row, column=3, value=num_or_dash(row["Dư nợ cho vay GDKQ (1)"]))
        ws.cell(row=excel_row, column=4, value=num_or_dash(row["Số lượng chứng khoán cho vay (2)"]))
        ws.cell(row=excel_row, column=5, value=num_or_dash(row["Số lượng chứng khoán niêm yết (3)"]))
        ws.cell(row=excel_row, column=6, value=num_or_dash(row["Vốn chủ sở hữu của CTCK (4)"]))

        v1 = row["Tỷ lệ dư nợ/VCSH (1)/(4)"]
        v2 = row["Tỷ lệ CK cho vay/CKNY (2)/(3)"]
        ws.cell(
            row=excel_row,
            column=7,
            value=float(v1) if pd.notna(v1) else None,
        )
        ws.cell(
            row=excel_row,
            column=8,
            value=float(v2) if pd.notna(v2) else None,
        )

    total_row = first_data_row + len(df)
    last_data_row = first_data_row + len(df) - 1

    total_debt = df["Dư nợ cho vay GDKQ (1)"].sum()
    total_qty = df["Số lượng chứng khoán cho vay (2)"].sum()

    ws.cell(row=total_row, column=2, value="Tổng cộng")
    ws.cell(row=total_row, column=3, value=num_or_dash(total_debt))
    ws.cell(row=total_row, column=4, value=num_or_dash(total_qty))

    title_font = Font(bold=True, size=16)
    unit_font = Font(italic=True, size=13)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"].font = unit_font
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for row in ws.iter_rows(min_row=3, max_row=last_data_row, min_col=1, max_col=n_cols):
        for cell in row:
            cell.border = border

    for col_idx in range(2, 5):
        ws.cell(row=total_row, column=col_idx).border = border

    bold_font = Font(bold=True)
    for col_idx in (3, 4):
        ws.cell(row=total_row, column=col_idx).font = bold_font

    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=1, max_col=n_cols):
        row[2].number_format = "#,##0"
        row[3].number_format = "#,##0"
        row[4].number_format = "#,##0"
        row[5].number_format = "#,##0"
        row[6].number_format = "0.00%"
        row[7].number_format = "0.00%"

    for row in ws.iter_rows(min_row=4, max_row=total_row, min_col=3, max_col=8):
        for cell in row:
            cell.alignment = right_align

    width_map = {
        1: 5,
        2: 25,
        3: 25,
        4: 25,
        5: 25,
        6: 25,
        7: 25,
        8: 25,
    }
    for col_idx, width in width_map.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A4"

    max_row = total_row
    max_col = n_cols
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                new_font = copy(cell.font)
                new_font.name = "Times New Roman"
                cell.font = new_font

    wb.save(output_report_path)
    print(f"Da tao file bao cao ngay tu DF: {output_report_path}")


def today_trade_date_str(tz_name: str = "Asia/Ho_Chi_Minh") -> str:
    """Lấy trade_date dạng YYYY-MM-DD theo múi giờ VN (nếu muốn dùng tay)."""
    tz = ZoneInfo(tz_name)
    return datetime.now(tz).date().isoformat()

def build_daily_output_dir(base_daily_dir: str, trade_date: str) -> str:
    """
    base_daily_dir: \\DATAANALYTICS\...\Daily_report
    trade_date: 'YYYY-MM-DD'
    return: \\...\Daily_report\YYYY\Tháng M\DD
    """
    d = datetime.strptime(trade_date, "%Y-%m-%d").date()
    year = f"{d.year}"
    month = f"Tháng {d.month}"
    day = f"{d.day:02d}"
    return os.path.join(base_daily_dir, year, month, day)

# =========================
# 3. Main
# =========================

if __name__ == "__main__":
    # === Config chung ===
    EQUITY_VND_MILLION = 2_077_000.0

    # True: đọc từ Data Warehouse
    # False: đọc từ file Excel raw như cũ
    USE_DW = True

    BASE_DIR = r"\\DATAANALYTICS\Risk Management\Report\[RMD] Report SSC_Mid"
    DAILY_DIR = r"\\DATAANALYTICS\Risk Management\Report\[RMD] Report SSC_Mid\Report_SSC"

    if USE_DW:
        # MODE 1: Data Warehouse + Excel special room
        engine = get_dw_engine()

        # Nếu có truyền ngày ở dòng lệnh thì ưu tiên dùng ngày đó
        # Cú pháp: python build_margin_report.py 2025-12-10
        if len(sys.argv) >= 2:
            TRADE_DATE = sys.argv[1]   # giả định đã ở dạng 'YYYY-MM-DD'
            print(f"TRADE_DATE dùng cho báo cáo (nhập tay): {TRADE_DATE}")
        else:
            TRADE_DATE = get_latest_trade_date(engine)  # 'YYYY-MM-DD'
            print(f"TRADE_DATE dùng cho báo cáo (NGÀY MỚI NHẤT): {TRADE_DATE}")

        SPECIAL_ROOM_FILE = os.path.join(BASE_DIR, "Data gửi Nam.xlsx")

        df_full, special_codes = build_margin_report_df_from_dw(
            engine=engine,
            trade_date=TRADE_DATE,
            equity_vnd_million=EQUITY_VND_MILLION,
            margin_table="[DWH-CoSo].[dbo].[vpr0109]",
            used_table="[DWH-CoSo].[dbo].[230007]",
            vol_table="[DWH-CoSo].[dbo].[Data_Hop_Infos]",
            price_table="[DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]",
            #special room từ DW
            special_room_table="[DWH-CoSo].[dbo].[vpr0108]",
            special_room_excel_path=None,
            #special room từ file excel
            # special_room_table=None,
            # special_room_excel_path=SPECIAL_ROOM_FILE,
        )

        output_dir = build_daily_output_dir(DAILY_DIR, TRADE_DATE)
        os.makedirs(output_dir, exist_ok=True)

        trade_date_obj = datetime.strptime(TRADE_DATE, "%Y-%m-%d").date()
        date_ddmmyyyy = trade_date_obj.strftime("%d%m%Y")     # 11122025
        date_label = trade_date_obj.strftime("%Y.%m.%d")      # 2025.12.11

        RAW_OUTPUT = os.path.join(
            output_dir,
            f"180426__RMD_SCMS_Bao cao ngay truoc 8AM {date_ddmmyyyy}.xlsx",
        )
        REPORT_FILE = os.path.join(
            output_dir,
            f"Báo cáo ngày {date_label}.xlsx",
        )

    else:
        # MODE 2: Excel RAW như cũ
        SPECIAL_ROOM_FILE = os.path.join(BASE_DIR, "Data gửi Nam.xlsx")
        RAW_FILE = os.path.join(BASE_DIR, "Book10-12.xlsx")

        df_full = build_margin_report_df(
            excel_path=RAW_FILE,
            equity_vnd_million=EQUITY_VND_MILLION,
            special_room_path=SPECIAL_ROOM_FILE,
        )

        special_codes = set(
            df_full.loc[df_full["Total Used room"].fillna(0) > 0, "code"]
            .astype(str)
            .str.strip()
        )

        run_date_iso = today_trade_date_str()  # 'YYYY-MM-DD'
        run_date_obj = datetime.strptime(run_date_iso, "%Y-%m-%d").date()
        date_ddmmyyyy = run_date_obj.strftime("%d%m%Y")
        date_label = run_date_obj.strftime("%Y.%m.%d")

        output_dir = build_daily_output_dir(DAILY_DIR, run_date_iso)
        os.makedirs(output_dir, exist_ok=True)
        

        RAW_OUTPUT = os.path.join(
            output_dir,
            f"180426__RMD_SCMS_Bao cao ngay truoc 8AM excel {date_ddmmyyyy}.xlsx",
        )
        REPORT_FILE = os.path.join(
            output_dir,
            f"Báo cáo ngày excel {date_label}.xlsx",
        )

    # Ghi file từ df_full (dùng chung cho cả 2 mode)
    write_margin_report_sheet_from_df(
        df=df_full,
        special_codes=special_codes,
        excel_output_path=RAW_OUTPUT,
        sheet_name="Dư nợ theo used room",
    )
    write_margin_report_file_from_df(
        df_full=df_full,
        output_report_path=REPORT_FILE,
        sheet_name="Report",
    )
