# -*- coding: utf-8 -*-
"""「一覧表.xlsx」の請求書シートを使って、全生徒分の請求書PDFを1人1ファイルずつ
一括生成するスクリプト。

Excel(win32com)を裏で操作し、請求書シートの生徒番号(No)を1人ずつ切り替えながら
印刷範囲（A1:F、請求書本体だけ）をPDFに書き出す。生成したPDFをGmailの作成画面に
手動で添付すれば、そのまま保護者への送付に使える。

必要環境: Windows + Microsoft Excel（インストール済みであること）
"""

import os

import win32com.client as win32

EXCEL_PATH = os.path.abspath("一覧表.xlsx")
OUTPUT_DIR = os.path.abspath("請求書PDF")

MAIN_SHEET = "月末請求一覧"
INVOICE_SHEET = "請求書"
NO_CELL = "J4"
DATA_START_ROW = 4  # 月末請求一覧シート：生徒データの開始行
NO_COL = 1  # A列
NAME_COL = 2  # B列

xlTypePDF = 0


def safe_filename(name):
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "_")
    return name


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # DispatchEx で必ず新しいExcelプロセスを起動する。Dispatch だと既に起動中の
    # Excel（利用者が別のブックを開いて作業中かもしれない）につながってしまい、
    # 最後の excel.Quit() でそのブックごと閉じてしまう事故につながるため。
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(EXCEL_PATH)
    try:
        main_ws = wb.Sheets(MAIN_SHEET)
        invoice_ws = wb.Sheets(INVOICE_SHEET)

        # 生徒一覧（No・名前）を、非表示になっている月末請求一覧シートから読み取る。
        students = []
        row = DATA_START_ROW
        while True:
            no = main_ws.Cells(row, NO_COL).Value
            if not isinstance(no, (int, float)):
                break
            name = main_ws.Cells(row, NAME_COL).Value
            students.append((int(no), str(name)))
            row += 1

        print(f"{len(students)}名分の請求書PDFを作成します…")

        for no, name in students:
            invoice_ws.Range(NO_CELL).Value = no
            excel.Calculate()

            filename = f"{no:03d}_{safe_filename(name)}.pdf"
            pdf_path = os.path.join(OUTPUT_DIR, filename)
            invoice_ws.ExportAsFixedFormat(xlTypePDF, pdf_path)
            print(f"  {filename}")

        # 選択状態を空に戻してから閉じる（保存はしない＝元のファイルは変更しない）。
        invoice_ws.Range(NO_CELL).Value = None
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    print(f"完了：{OUTPUT_DIR} に請求書PDFを{len(students)}件出力しました。")


if __name__ == "__main__":
    main()
