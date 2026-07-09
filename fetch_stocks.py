"""
日本株 株価自動取得スクリプト
- スプレッドシートのA列（証券コードの数字）とB列（接尾辞、例: T）を読み、
  yfinance 用に「1234.T」のように結合する
- 終値をH列、前日比(%)をI列、1か月前比(%)をJ列、3か月前比(%)をK列に上書きする
- 決算発表日をN列に上書きする
"""

import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
import math
import time
import os
import json

# ============================================================
# Google Sheets 設定
# ============================================================
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SHEET_NAME     = "最新株価"
JST = ZoneInfo("Asia/Tokyo")

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


def code_from_cell(a_raw) -> str:
    """A列のセル値を証券コードの数字部分に正規化する（シートの数値・文字列の両方に対応）"""
    if a_raw is None:
        return ""
    s = str(a_raw).strip()
    if not s:
        return ""
    try:
        n = float(s.replace(",", ""))
        if not math.isfinite(n):
            return s
        if abs(n - round(n)) < 1e-9:
            return str(int(round(n)))
        return s
    except ValueError:
        return s


def parse_row_to_ticker(row: list) -> str | None:
    """1行分の A,B からティッカー文字列を組み立てる。不足時は None"""
    padded = (list(row) + ["", ""])[:2]
    a_raw, b_raw = padded[0], padded[1]
    code = code_from_cell(a_raw)
    suffix = str(b_raw).strip() if b_raw is not None else ""
    if not code or not suffix:
        return None
    return f"{code}.{suffix}"


def _close_near_days_ago(hist, days: int) -> float | None:
    """履歴からおおよそ days 日前の終値を返す"""
    if hist.empty:
        return None
    target = hist.index[-1] - timedelta(days=days)
    past = hist.loc[hist.index <= target]
    if past.empty:
        return None
    try:
        v = float(past.iloc[-1]["Close"])
        return v if math.isfinite(v) else None
    except (TypeError, ValueError, KeyError):
        return None


def _pct_change(current: float, past: float | None) -> float | str:
    if past is None or past == 0:
        return ""
    return round((current - past) / past * 100, 2)


def _as_date(value) -> date | None:
    """yfinance の日付値を date に正規化する"""
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if hasattr(value, "to_pydatetime"):
        try:
            return value.to_pydatetime().date()
        except (TypeError, ValueError, AttributeError):
            return None
    if hasattr(value, "date"):
        try:
            d = value.date()
            return d if isinstance(d, date) else None
        except (TypeError, ValueError, AttributeError):
            return None
    return None


def _format_earnings_date(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def _fetch_earnings_date(stock) -> str:
    """次回決算発表日を YYYY-MM-DD 形式で返す。取得不可なら空文字"""
    today = datetime.now(JST).date()

    try:
        cal = stock.calendar
        if isinstance(cal, dict):
            raw_dates = cal.get("Earnings Date")
            if raw_dates is not None:
                if not isinstance(raw_dates, list):
                    raw_dates = [raw_dates]
                parsed = [_as_date(d) for d in raw_dates]
                parsed = [d for d in parsed if d is not None]
                future = sorted(d for d in parsed if d >= today)
                if future:
                    return _format_earnings_date(future[0])
                if parsed:
                    return _format_earnings_date(min(parsed))
    except Exception:
        pass

    try:
        earnings_dates = stock.earnings_dates
        if earnings_dates is not None and not earnings_dates.empty:
            for idx, row in earnings_dates.iterrows():
                d = _as_date(idx)
                if d is None:
                    continue
                reported = row.get("Reported EPS") if hasattr(row, "get") else None
                if reported is None or (isinstance(reported, float) and math.isnan(reported)):
                    if d >= today:
                        return _format_earnings_date(d)
            for idx in earnings_dates.index:
                d = _as_date(idx)
                if d is not None and d >= today:
                    return _format_earnings_date(d)
    except Exception:
        pass

    try:
        info = stock.info or {}
        for key in ("earningsTimestampStart", "earningsTimestamp", "earningsTimestampEnd"):
            ts = info.get(key)
            if ts is None:
                continue
            d = datetime.fromtimestamp(ts, tz=JST).date()
            if d >= today:
                return _format_earnings_date(d)
    except Exception:
        pass

    return ""


def fetch_one(ticker: str):
    """1銘柄の終値、前日比(%)、1か月前比(%)、3か月前比(%)、決算発表日を取得して返す"""
    try:
        stock = yf.Ticker(ticker)
        hist  = stock.history(period="6mo")

        earnings_date = _fetch_earnings_date(stock)

        if hist.empty:
            print(f"  WARNING {ticker}: データなし")
            return None, None, None, None, earnings_date

        latest = hist.iloc[-1]
        prev   = hist.iloc[-2] if len(hist) >= 2 else None

        try:
            close_f = float(latest["Close"])
            if not math.isfinite(close_f):
                raise ValueError
        except (TypeError, ValueError):
            print(f"  WARNING {ticker}: 終値が無効な値")
            return None, None, None, None, earnings_date

        close = round(close_f, 1)

        prev_close = None
        if prev is not None:
            try:
                prev_f = float(prev["Close"])
                if math.isfinite(prev_f):
                    prev_close = prev_f
            except (TypeError, ValueError):
                pass

        change_pct  = _pct_change(close_f, prev_close)
        change_1m   = _pct_change(close_f, _close_near_days_ago(hist, 30))
        change_3m   = _pct_change(close_f, _close_near_days_ago(hist, 90))

        return close, change_pct, change_1m, change_3m, earnings_date

    except Exception as e:
        print(f"  ERROR {ticker}: {e}")
        return None, None, None, None, ""


# ============================================================
# メイン
# ============================================================

def main():
    updated_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")
    print(f"\n{'='*50}")
    print(f"  株価取得開始: {updated_at}")
    print(f"{'='*50}\n")

    # --- Google Sheets に接続 ---
    client = get_gspread_client()
    ss     = client.open_by_key(SPREADSHEET_ID)
    ws     = ss.worksheet(SHEET_NAME)

    # --- A列・B列からティッカーを組み立て（2行目以降、範囲 A2:B）---
    rows = ws.get_values("A2:B")
    if not rows:
        print("  データ行がありません（A2:B が空）。")
        return

    valid_total = sum(1 for row in rows if parse_row_to_ticker(row) is not None)
    print(f"  スプレッドシートから {len(rows)} 行を読み込み（うち {valid_total} 銘柄を取得）\n")

    # --- 株価・決算発表日取得 ---
    hijk_values: list[list] = []
    n_values: list[list] = []
    done = 0

    for row in rows:
        ticker = parse_row_to_ticker(row)
        if ticker is None:
            hijk_values.append(["", "", "", ""])
            n_values.append([""])
            continue
        done += 1
        print(f"  [{done:>3}/{valid_total}] {ticker} 取得中...")
        close, change_pct, change_1m, change_3m, earnings_date = fetch_one(ticker)
        hijk_values.append(
            [
                close if close is not None else "",
                change_pct if change_pct is not None else "",
                change_1m if change_1m is not None else "",
                change_3m if change_3m is not None else "",
            ]
        )
        n_values.append([earnings_date])
        time.sleep(0.3)

    # --- H〜K列・N列に一括書き込み ---
    print("\n  スプレッドシートに書き込み中...")
    last_row = len(rows) + 1  # 2行目スタートなので +1

    ws.update(range_name=f"H2:K{last_row}", values=hijk_values, value_input_option="USER_ENTERED")
    ws.update(range_name=f"N2:N{last_row}", values=n_values, value_input_option="USER_ENTERED")
    ws.update(range_name="O1", values=[[updated_at]], value_input_option="USER_ENTERED")

    print(f"\n{'='*50}")
    print(f"  完了！ {valid_total} 銘柄を更新しました（全 {len(rows)} 行）")
    print(f"  更新日時: {updated_at}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()