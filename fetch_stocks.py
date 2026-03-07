"""
日本株 株価自動取得スクリプト
毎日GitHub Actionsで実行 → Google スプレッドシートに書き込み
"""

import yfinance as yf
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time
import os
import json

# ============================================================
# 設定: 取得する銘柄コードリスト（300社分をここに記載）
# ============================================================
TICKERS = [
    "7203.T",  # トヨタ自動車
    "6758.T",  # ソニーグループ
    "9984.T",  # ソフトバンクグループ
    "8306.T",  # 三菱UFJフィナンシャル・グループ
    "6861.T",  # キーエンス
    "9432.T",  # NTT
    "8035.T",  # 東京エレクトロン
    "4063.T",  # 信越化学工業
    "7741.T",  # HOYA
    "6954.T",  # ファナック
    # ↓ここに残りの銘柄コードを追加（例: "1234.T",）
]

# ============================================================
# Google Sheets 設定
# ============================================================
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")  # GitHub Secretsから取得
SHEET_LATEST   = "最新株価"   # 最新データを上書きするシート
SHEET_HISTORY  = "履歴"       # 日次で追記していくシート

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
    """1銘柄の株価データを取得して辞書で返す"""
    try:
        stock = yf.Ticker(ticker)
        hist  = stock.history(period="2d")  # 前日比のため2日分取得

        if hist.empty:
            print(f"  WARNING  {ticker}: データなし")
            return None

        latest     = hist.iloc[-1]
        prev       = hist.iloc[-2] if len(hist) >= 2 else None
        close      = latest["Close"]
        prev_close = prev["Close"] if prev is not None else None
        change     = round(close - prev_close, 1)                        if prev_close else ""
        change_pct = round((close - prev_close) / prev_close * 100, 2)  if prev_close else ""

        # 銘柄名（取得失敗してもスキップ）
        try:
            info = stock.info
            name = info.get("longName") or info.get("shortName") or ticker
        except Exception:
            name = ticker

        return {
            "ticker":     ticker,
            "name":       name,
            "date":       latest.name.strftime("%Y-%m-%d"),
            "open":       round(latest["Open"],  1),
            "high":       round(latest["High"],  1),
            "low":        round(latest["Low"],   1),
            "close":      round(close,           1),
            "volume":     int(latest["Volume"]),
            "prev_close": round(prev_close, 1)   if prev_close else "",
            "change":     change,
            "change_pct": change_pct,
        }
    except Exception as e:
        print(f"  ERROR {ticker}: {e}")
        return None


def to_rows(data_list):
    """辞書リストをスプレッドシート書き込み用の2次元リストに変換"""
    return [
        [
            d["ticker"], d["name"], d["date"],
            d["open"], d["high"], d["low"], d["close"],
            d["volume"], d["prev_close"], d["change"], d["change_pct"],
        ]
        for d in data_list
    ]


def update_latest_sheet(ws, data_list, updated_at):
    """「最新株価」シートを全件上書き"""
    headers = [
        "銘柄コード", "銘柄名", "取得日",
        "始値", "高値", "安値", "終値",
        "出来高", "前日終値", "前日比(円)", "前日比(%)",
    ]
    ws.clear()
    ws.update("A1", [headers] + to_rows(data_list))
    ws.format("A1:K1", {
        "textFormat": {"bold": True},
        "backgroundColor": {"red": 0.20, "green": 0.40, "blue": 0.75},
        "horizontalAlignment": "CENTER",
    })
    ws.update("M1", [[f"最終更新: {updated_at}"]])
    print(f"  OK 「{ws.title}」シート更新完了 ({len(data_list)} 銘柄)")


def append_history_sheet(ws, data_list):
    """「履歴」シートに今日分を追記（ヘッダーがなければ先に挿入）"""
    existing = ws.get_all_values()
    if not existing:
        headers = [
            "銘柄コード", "銘柄名", "取得日",
            "始値", "高値", "安値", "終値",
            "出来高", "前日終値", "前日比(円)", "前日比(%)",
        ]
        ws.append_row(headers)
        ws.format("A1:K1", {"textFormat": {"bold": True}})

    ws.append_rows(to_rows(data_list), value_input_option="USER_ENTERED")
    print(f"  OK 「{ws.title}」シートに {len(data_list)} 行追記")


# ============================================================
# メイン
# ============================================================

def main():
    updated_at = datetime.now().strftime("%Y-%m-%d %H:%M JST")
    print(f"\n{'='*50}")
    print(f"  株価取得開始: {updated_at}")
    print(f"  対象銘柄数: {len(TICKERS)}")
    print(f"{'='*50}\n")

    # --- 株価取得 ---
    data_list = []
    for i, ticker in enumerate(TICKERS, 1):
        print(f"  [{i:>3}/{len(TICKERS)}] {ticker} 取得中...")
        result = fetch_one(ticker)
        if result:
            data_list.append(result)
        time.sleep(0.5)  # API負荷軽減のため少し待機

    print(f"\n  取得成功: {len(data_list)} / {len(TICKERS)} 銘柄\n")

    if not data_list:
        print("  WARNING 取得データが0件のため終了します")
        return

    # --- Google Sheets に書き込み ---
    print("  Google Sheets に書き込み中...")
    client = get_gspread_client()
    ss     = client.open_by_key(SPREADSHEET_ID)

    # シートが存在しない場合は自動作成
    sheet_titles = [ws.title for ws in ss.worksheets()]
    if SHEET_LATEST  not in sheet_titles:
        ss.add_worksheet(title=SHEET_LATEST,  rows=400,    cols=15)
    if SHEET_HISTORY not in sheet_titles:
        ss.add_worksheet(title=SHEET_HISTORY, rows=100000, cols=15)

    update_latest_sheet( ss.worksheet(SHEET_LATEST),  data_list, updated_at)
    append_history_sheet(ss.worksheet(SHEET_HISTORY), data_list)

    print(f"\n{'='*50}")
    print(f"  完了！")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
