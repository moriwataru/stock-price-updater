"""
日本株 株価自動取得スクリプト
- スプレッドシートのB列から証券コードを読み取る
- 終値をF列、前日比(%)をG列に上書きする
"""

import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import math
import time
import os
import json

# ============================================================
# Google Sheets 設定
# ============================================================
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SHEET_NAME     = "最新株価"

# ============================================================
# 関数定義
# ============================================================

def get_gspread_client():
    """Google Sheets クライアントを返す"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json:
        raise ValueError("環境変数 GOOGLE_CREDENTIALS_JSON が未設定です")

    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def fetch_one(ticker: str):
    """1銘柄の終値と前日比(%)を取得して返す"""
    try:
        stock = yf.Ticker(ticker)
        hist  = stock.history(period="2d")

        if hist.empty:
            print(f"  WARNING {ticker}: データなし")
            return None, None

        latest    = hist.iloc[-1]
        prev      = hist.iloc[-2] if len(hist) >= 2 else None
        close_raw = latest["Close"]

        if close_raw is None or (isinstance(close_raw, float) and not math.isfinite(close_raw)):
            print(f"  WARNING {ticker}: 終値が無効な値")
            return None, None

        close      = round(float(close_raw), 1)
        prev_close = prev["Close"] if prev is not None else None

        if prev_close is not None and (not isinstance(prev_close, (int, float)) or not math.isfinite(float(prev_close))):
            prev_close = None

        change_pct = round((close - float(prev_close)) / float(prev_close) * 100, 2) if prev_close else ""

        return close, change_pct

    except Exception as e:
        print(f"  ERROR {ticker}: {e}")
        return None, None


# ============================================================
# メイン
# ============================================================

def main():
    updated_at = datetime.now().strftime("%Y-%m-%d %H:%M JST")
    print(f"\n{'='*50}")
    print(f"  株価取得開始: {updated_at}")
    print(f"{'='*50}\n")

    # --- Google Sheets に接続 ---
    client = get_gspread_client()
    ss     = client.open_by_key(SPREADSHEET_ID)
    ws     = ss.worksheet(SHEET_NAME)

    # --- B列から証券コードを取得（2行目以降）---
    b_col   = ws.col_values(2)
    tickers = [v.strip() for v in b_col[1:] if v.strip()]  # 2行目以降・空白除外
    print(f"  スプレッドシートから {len(tickers)} 銘柄を読み込みました\n")

    # --- 株価取得 ---
    close_col      = []  # F列に書き込む終値
    change_pct_col = []  # G列に書き込む前日比(%)

    for i, ticker in enumerate(tickers, 1):
        print(f"  [{i:>3}/{len(tickers)}] {ticker} 取得中...")
        close, change_pct = fetch_one(ticker)
        close_col.append([close if close is not None else ""])
        change_pct_col.append([change_pct if change_pct is not None else ""])
        time.sleep(0.3)

    # --- F列・G列に一括書き込み ---
    print("\n  スプレッドシートに書き込み中...")
    last_row = len(tickers) + 1  # 2行目スタートなので+1

    ws.update(range_name=f"G2:G{last_row}", values=close_col,      value_input_option="USER_ENTERED")
    ws.update(range_name=f"H2:H{last_row}", values=change_pct_col, value_input_option="USER_ENTERED")

    print(f"\n{'='*50}")
    print(f"  完了！ {len(tickers)} 銘柄を更新しました")
    print(f"  更新日時: {updated_at}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()