#!/usr/bin/env python3
"""
営業日報ExcelをGoogleスプレッドシートに転記するスクリプト

使い方:
    python main.py <Excelファイルパス> <年> <月>

例:
    python main.py 高円寺_日計報告書.xlsx 2026 3
"""

import sys
from parse_excel import parse_daily_report
from update_sheets import update_spreadsheet


def main():
    if len(sys.argv) < 4:
        print("=" * 50)
        print("使い方: python main.py <Excelファイルパス> <年> <月>")
        print("例:     python main.py 高円寺_日計報告書.xlsx 2026 3")
        print("=" * 50)
        sys.exit(1)

    excel_path = sys.argv[1]
    year  = int(sys.argv[2])
    month = int(sys.argv[3])

    print("=" * 50)
    print(f"ファイル: {excel_path}")
    print(f"対象月:   {year}年{month}月")
    print("=" * 50)

    # ── Excelを読み込む ──────────────────────────
    print("\n【1】Excelを読み込み中...")
    data = parse_daily_report(excel_path)

    print(f"  店舗: {data['store']}")
    print(f"  スタイリスト数: {len(data['stylists'])}人")
    for s in data['stylists']:
        売上   = f"{int(s['売上']):,}"   if s.get('売上')   else '-'
        生産性 = f"{int(s['生産性']):,}" if s.get('生産性') else '-'
        稼働   = f"{s['稼働時間']}"      if s.get('稼働時間') else '-'
        単価   = f"{int(s['客単価']):,}" if s.get('客単価') else '-'
        空き枠率 = f"{s['空き枠率']:.1f}%" if s.get('空き枠率') is not None else '-'
        print(f"    {s['name']}: 売上={売上}円 / 生産性={生産性} / 稼働時間={稼働}h / 客単価={単価}円 / 空き枠率={空き枠率}")

    # ── スプレッドシートを更新 ───────────────────
    print(f"\n【2】Googleスプレッドシートを更新中...")
    update_spreadsheet(data, year, month)

    print("\n" + "=" * 50)
    print("完了！")
    print("=" * 50)


if __name__ == '__main__':
    main()
