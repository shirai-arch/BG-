"""
Googleスプレッドシートを更新するモジュール
"""
import gspread
from google.oauth2.service_account import Credentials
from config import SPREADSHEET_ID, SHEET_GID, CREDENTIALS_PATH

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

MONTH_NAMES = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月']

# スプレッドシートの各スタイリストの行構成
# 名前行 → 生産性、名前行+1 → 売上、名前行+2 → 稼働時間、名前行+3 → 客単価
ROW_OFFSETS = {
    '生産性':  0,
    '売上':    1,
    '稼働時間': 2,
    '客単価':  3,
}


def get_sheet():
    """認証してワークシートを返す"""
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    for ws in spreadsheet.worksheets():
        if ws.id == SHEET_GID:
            return ws
    raise ValueError(f"シートが見つかりません (gid={SHEET_GID})")


def find_month_col(all_values, year, month):
    """
    ヘッダー行を読んで、指定した年・月の「実績」列インデックス(0-indexed)を返す。
    スプレッドシートの列構成（月ごと）: 目標 | 2025実績 | 2026実績 | 成長率
    """
    month_name = MONTH_NAMES[month - 1]   # 例: "3月"
    year_label = f"{year}実績"             # 例: "2026実績"

    month_row_idx = None
    month_col_start = None

    # 月名がある行を探す
    for row_idx, row in enumerate(all_values):
        for col_idx, cell in enumerate(row):
            if str(cell).strip() == month_name:
                month_row_idx = row_idx
                month_col_start = col_idx
                break
        if month_row_idx is not None:
            break

    if month_row_idx is None:
        return None

    # 月名の次の行で year_label を探す（月名列から最大5列以内）
    next_row = all_values[month_row_idx + 1] if month_row_idx + 1 < len(all_values) else []
    for col_idx in range(month_col_start, min(month_col_start + 5, len(next_row))):
        if year_label in str(next_row[col_idx]):
            return col_idx  # 0-indexed

    return None


def find_stylist_row(all_values, stylist_name):
    """
    スタイリスト名でシート全体を検索し、名前が入っている行番号(0-indexed)を返す。
    姓のみで一致する場合も考慮する（例: シート「望月」← Excel「望月　智博」）
    """
    # 姓（スペース前の部分）を取得
    last_name = stylist_name.replace('　', ' ').split()[0]

    for row_idx, row in enumerate(all_values):
        for cell in row:
            cell_str = str(cell).replace('　', ' ').strip()
            if not cell_str:
                continue
            # 完全一致 または 姓のみ一致
            if cell_str == stylist_name or cell_str == last_name:
                return row_idx

    return None


def update_spreadsheet(data, year, month):
    """
    Googleスプレッドシートに日計報告書のデータを書き込む

    Args:
        data: parse_daily_report() の戻り値
        year: 年 (例: 2026)
        month: 月 (例: 3)
    """
    print("  Googleスプレッドシートに接続中...")
    sheet = get_sheet()
    all_values = sheet.get_all_values()

    # 対象月の列インデックスを取得
    col_idx = find_month_col(all_values, year, month)
    if col_idx is None:
        raise ValueError(f"{year}年{month}月の列がスプレッドシートで見つかりません")

    print(f"  {year}年{month}月の列: {col_idx + 1}列目 ({_col_letter(col_idx + 1)}列)")

    updates = []
    not_found = []

    for stylist in data['stylists']:
        name = stylist['name']
        row_idx = find_stylist_row(all_values, name)

        if row_idx is None:
            not_found.append(name)
            continue

        print(f"  ✓ {name} → {row_idx + 1}行目")

        # 4指標（生産性・売上・稼働時間・客単価）を書き込む
        metrics = {
            '生産性':  stylist.get('生産性'),
            '売上':    stylist.get('売上'),
            '稼働時間': stylist.get('稼働時間'),
            '客単価':  stylist.get('客単価'),
        }

        for metric, offset in ROW_OFFSETS.items():
            val = metrics.get(metric)
            if val is None:
                continue
            target_row = row_idx + offset + 1  # 1-indexed
            target_col = col_idx + 1           # 1-indexed
            cell_a1 = gspread.utils.rowcol_to_a1(target_row, target_col)
            updates.append({'range': cell_a1, 'values': [[val]]})

    # バッチ更新（APIコールを1回にまとめる）
    if updates:
        sheet.batch_update(updates)
        print(f"\n  → {len(updates)} セルを更新しました")
    else:
        print("\n  → 更新するデータがありませんでした")

    if not_found:
        print(f"\n  ⚠ 以下のスタイリストはシートで見つかりませんでした:")
        for name in not_found:
            print(f"    - {name}")


def _col_letter(col_num):
    """列番号(1-indexed)をアルファベット表記に変換 (例: 1→A, 27→AA)"""
    result = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result
