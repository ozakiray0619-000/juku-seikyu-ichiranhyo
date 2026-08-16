# -*- coding: utf-8 -*-
"""塾の月末請求一覧表を生成するスクリプト。

openpyxl を使ってダミー生徒180名分の請求データを作成し、
「一覧表.xlsx」として保存する。
"""

import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

random.seed(42)

# ------------------------------------------------------------
# ダミーデータ生成
# ------------------------------------------------------------

SEI_LIST = [
    "佐藤", "鈴木", "高橋", "田中", "伊藤", "渡辺", "山本", "中村", "小林", "加藤",
    "吉田", "山田", "佐々木", "山口", "松本", "井上", "木村", "林", "斎藤", "清水",
    "山崎", "森", "池田", "橋本", "阿部", "石川", "山下", "中島", "石井", "小川",
    "前田", "岡田", "長谷川", "藤田", "後藤", "近藤", "村上", "遠藤", "青木", "坂本",
    "福田", "太田", "西村", "藤井", "岡本", "中野", "小野", "田村", "竹内", "金子",
]

MEI_LIST = [
    "翔太", "陽菜", "大輝", "美咲", "健太", "愛", "拓也", "結衣", "颯太", "さくら",
    "蓮", "陽菜乃", "悠斗", "葵", "大和", "美月", "湊", "花音", "樹", "莉子",
    "陸", "楓", "颯", "凛", "駿", "詩織", "航", "彩花", "翼", "優奈",
    "尊", "咲良", "結翔", "心春", "陽向", "美結", "朝陽", "琴音", "悠真", "千尋",
    "一輝", "美桜", "大翔", "萌", "海斗", "七海", "蒼", "遥", "煌", "紬",
]

CLASSROOMS = [
    ("本校", 70),
    ("東校", 60),
    ("西校", 50),
]

GRADES = [
    "小4", "小5", "小6",
    "中1", "中2", "中3",
    "高1", "高2", "高3",
]

# コース別の単価マスター: (コース名, 1コマ単価, 標準教材費)
# 「単価表」シートにテーブルとして書き出し、一覧表の1コマ単価・教材費はここから
# 自動で参照する（コースが決まれば単価・教材費が自動的に決まる）。
COURSE_MASTER = [
    ("個別指導S", 4000, 7000),
    ("個別指導", 3500, 6000),
    ("集団授業", 2500, 4500),
    ("映像授業", 2100, 3000),
]
COURSE_NAMES = [c[0] for c in COURSE_MASTER]


def make_unique_names(count):
    """重複しない生徒名（姓+名）を count 件生成する。"""
    names = set()
    while len(names) < count:
        name = random.choice(SEI_LIST) + random.choice(MEI_LIST)
        names.add(name)
    names = list(names)
    random.shuffle(names)
    return names


def build_students():
    total = sum(n for _, n in CLASSROOMS)
    all_names = make_unique_names(total)

    students = []
    name_idx = 0
    for classroom, count in CLASSROOMS:
        for _ in range(count):
            name = all_names[name_idx]
            name_idx += 1

            grade = random.choice(GRADES)
            course = random.choice(COURSE_NAMES)
            # 週あたりの通塾回数（実際の登録データに相当する、人が入力する値）
            weekly_lessons = random.randint(1, 5)

            students.append(
                {
                    "name": name,
                    "classroom": classroom,
                    "grade": grade,
                    "course": course,
                    "weekly_lessons": weekly_lessons,
                    "email": f"seito{len(students) + 1:03d}@example.com",
                }
            )
    return students


# ------------------------------------------------------------
# Excel 出力
# ------------------------------------------------------------

HEADERS = [
    "No", "生徒名", "教室", "学年", "コース",
    "1コマ単価", "週コマ数", "コマ数", "教材費", "授業料", "請求額(税込)", "保護者メール",
]

MONEY_FORMAT = "#,##0"
TABLE_NAME = "SeikyuData"
TANKA_TABLE_NAME = "TankaData"

TOTAL_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
SUMMARY_FILL = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")
THIN = Side(style="thin", color="FFB0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_workbook(students):
    wb = Workbook()
    # openpyxl は数式の計算結果をキャッシュしないため、開いた瞬間に必ず
    # 再計算されるようにする(古い/空の値のまま表示されるのを防ぐ)。
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ws = wb.active
    ws.title = "月末請求一覧"

    build_tanka_sheet(wb)

    n_students = len(students)
    header_row = 3
    data_start_row = header_row + 1
    data_end_row = data_start_row + n_students - 1
    total_row = data_end_row + 1
    n_cols = len(HEADERS)

    # --- 合計金額（表のいちばん上、構造化参照で常に最新のテーブル範囲を集計） ---
    invoice_header = "請求額(税込)"
    ws.cell(row=1, column=1, value="合計金額（請求額の合計）")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    total_cell = ws.cell(
        row=1,
        column=2,
        value=f"=SUM({TABLE_NAME}[{invoice_header}])",
    )
    total_cell.number_format = MONEY_FORMAT
    total_cell.font = Font(bold=True, size=12, color="FFC00000")
    for col in range(1, n_cols + 1):
        ws.cell(row=1, column=col).fill = SUMMARY_FILL
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=4)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 6  # 空行（区切り）

    # --- 見出し行（配色・帯縞・フィルターはテーブルスタイルが担当） ---
    for col, title in enumerate(HEADERS, start=1):
        cell = ws.cell(row=header_row, column=col, value=title)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # --- データ行 ---
    # 列: A=No B=生徒名 C=教室 D=学年 E=コース F=1コマ単価 G=週コマ数
    #     H=コマ数 I=教材費 J=授業料 K=請求額(税込) L=保護者メール
    for i, student in enumerate(students):
        row = data_start_row + i
        # No は行位置から自動計算する数式にして、行の挿入/削除に追随させる
        ws.cell(row=row, column=1, value=f"=ROW()-{header_row}")
        ws.cell(row=row, column=2, value=student["name"])
        ws.cell(row=row, column=3, value=student["classroom"])
        ws.cell(row=row, column=4, value=student["grade"])
        ws.cell(row=row, column=5, value=student["course"])

        # 1コマ単価・教材費は「単価表」シートからコースに応じて自動参照する
        c_unit = ws.cell(
            row=row,
            column=6,
            value=(
                f'=IFERROR(INDEX({TANKA_TABLE_NAME}[1コマ単価],'
                f'MATCH(E{row},{TANKA_TABLE_NAME}[コース],0)),0)'
            ),
        )
        c_weekly = ws.cell(row=row, column=7, value=student["weekly_lessons"])
        # コマ数(月間)は「週コマ数×4週」で自動計算
        c_lessons = ws.cell(row=row, column=8, value=f"=G{row}*4")
        c_material = ws.cell(
            row=row,
            column=9,
            value=(
                f'=IFERROR(INDEX({TANKA_TABLE_NAME}[教材費(標準)],'
                f'MATCH(E{row},{TANKA_TABLE_NAME}[コース],0)),0)'
            ),
        )

        c_tuition = ws.cell(row=row, column=10, value=f"=F{row}*H{row}")
        c_invoice = ws.cell(
            row=row,
            column=11,
            value=f"=J{row}+I{row}+ROUND(I{row}*0.1,0)",
        )

        ws.cell(row=row, column=12, value=student["email"])

        for c in (c_unit, c_material, c_tuition, c_invoice):
            c.number_format = MONEY_FORMAT
        c_weekly.number_format = "0"
        c_lessons.number_format = "0"

        for col in (1, 3, 4, 5):
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")

    # --- 見出し+データ範囲を Excel テーブル（ListObject）化 ---
    table_ref = f"A{header_row}:{get_column_letter(n_cols)}{data_end_row}"
    table_columns = [
        TableColumn(id=idx, name=title) for idx, title in enumerate(HEADERS, start=1)
    ]
    tab = Table(displayName=TABLE_NAME, ref=table_ref, tableColumns=table_columns)
    tab.autoFilter = AutoFilter(ref=table_ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # --- 合計行（表のいちばん下、構造化参照で行の増減に自動追随） ---
    ws.cell(row=total_row, column=1, value="合計")
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="center")

    for header in ("コマ数", "教材費", "授業料", "請求額(税込)"):
        col = HEADERS.index(header) + 1
        cell = ws.cell(
            row=total_row,
            column=col,
            value=f"=SUM({TABLE_NAME}[{header}])",
        )
        cell.number_format = MONEY_FORMAT
        cell.font = Font(bold=True)

    for col in range(1, n_cols + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.fill = TOTAL_FILL
        cell.border = BORDER

    # --- 列幅 ---
    widths = [6, 14, 8, 6, 12, 11, 9, 8, 10, 12, 14, 22]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # --- 見出し行を固定表示 ---
    ws.freeze_panes = f"A{data_start_row}"

    build_invoice_sheet(wb, ws.title, data_start_row, data_end_row)

    return wb


def build_tanka_sheet(wb):
    """コース別の単価マスターシート。ここを直せば一覧表の該当コース全員に自動反映される。"""
    ws = wb.create_sheet("単価表")
    ws.sheet_view.showGridLines = False

    ws.cell(row=1, column=1, value="コース別 単価マスター").font = Font(
        bold=True, size=14, color="FF223757"
    )
    ws.cell(
        row=2,
        column=1,
        value="ここでコースの単価・教材費を変更すると、一覧表の該当コースの生徒全員に自動反映されます。",
    ).font = Font(size=10, color="FF5B6472")

    tanka_headers = ["コース", "1コマ単価", "教材費(標準)"]
    header_row = 4
    for col, title in enumerate(tanka_headers, start=1):
        ws.cell(row=header_row, column=col, value=title)

    for i, (course, unit_price, material_fee) in enumerate(COURSE_MASTER):
        row = header_row + 1 + i
        ws.cell(row=row, column=1, value=course)
        c_unit = ws.cell(row=row, column=2, value=unit_price)
        c_material = ws.cell(row=row, column=3, value=material_fee)
        c_unit.number_format = MONEY_FORMAT
        c_material.number_format = MONEY_FORMAT

    data_end_row = header_row + len(COURSE_MASTER)
    table_ref = f"A{header_row}:C{data_end_row}"
    table_columns = [
        TableColumn(id=idx, name=title) for idx, title in enumerate(tanka_headers, start=1)
    ]
    tab = Table(displayName=TANKA_TABLE_NAME, ref=table_ref, tableColumns=table_columns)
    tab.autoFilter = AutoFilter(ref=table_ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium7",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    for col, width in zip("ABC", (16, 14, 16)):
        ws.column_dimensions[col].width = width


def build_invoice_sheet(wb, main_sheet_name, data_start_row, data_end_row):
    """生徒名を選ぶとその生徒の個人請求書とメール送信用リンクが自動生成されるシート。"""
    ws = wb.create_sheet("請求書")
    ws.sheet_view.showGridLines = False

    # データ入力規則（ドロップダウン・数値範囲）は、別シートの「テーブルの構造化参照」
    # （例: SeikyuData[生徒名]）を直接読み込めない（Excelの既知の制限）。
    # 必ずプレーンなセル範囲参照（例: 月末請求一覧!$B$4:$B$183）を使う。
    name_range_ref = f"'{main_sheet_name}'!$B${data_start_row}:$B${data_end_row}"

    JUKU_NAME = "さくら学習塾"
    NO_CELL = "C4"
    NO_CELL_ABS = "$C$4"
    NAME_INPUT_CELL = "C5"
    NAME_INPUT_ABS = "$C$5"
    RESOLVED_CELL_ABS = "$C$6"  # No入力・名前入力のどちらからでも解決される選択中の生徒名

    def lookup(header):
        return f'IFERROR(INDEX({TABLE_NAME}[{header}],MATCH({RESOLVED_CELL_ABS},{TABLE_NAME}[生徒名],0)),"")'

    title = ws.cell(row=1, column=1, value="請求書")
    title.font = Font(bold=True, size=20, color="FF223757")
    ws.merge_cells("A1:D1")

    ws.cell(row=1, column=6, value="発行日：")
    ws.cell(row=1, column=6).alignment = Alignment(horizontal="right")
    issue_date = ws.cell(row=1, column=7, value="=TODAY()")
    issue_date.number_format = "yyyy/mm/dd"

    ws.cell(row=2, column=1, value=JUKU_NAME).font = Font(size=11, color="FF5B6472")

    # クリックの動作確認用に、生徒選択に関係なく常に表示される固定のGmailリンクを用意する
    # (openpyxl のネイティブハイパーリンク機能。数式のHYPERLINK関数とは別の仕組みで、
    # こちらが確実に動くかどうかで問題の切り分けができる)
    test_link_cell = ws.cell(row=3, column=1, value="✅ クリックテスト：Gmailを開く（動作確認用）")
    test_link_cell.hyperlink = "https://mail.google.com/mail/?view=cm&fs=1&su=クリックテスト"
    test_link_cell.font = Font(bold=True, color="FF1155CC", underline="single")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=4)

    thick_side = Side(style="medium", color="FF223757")
    input_fill = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
    input_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)

    # --- ① 生徒番号(No)で入力 ---
    ws.cell(row=4, column=1, value="① 生徒番号(No)で入力 ▶").font = Font(bold=True)
    no_cell = ws.cell(row=4, column=3, value=None)
    no_cell.font = Font(bold=True, size=14)
    no_cell.fill = input_fill
    no_cell.border = input_border
    no_cell.alignment = Alignment(horizontal="center")
    ws.merge_cells("C4:D4")

    dv_no = DataValidation(
        type="whole",
        operator="between",
        formula1=1,
        formula2=f"=COUNTA({name_range_ref})",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="入力エラー",
        error="一覧表にあるNo（生徒番号）を入力してください。",
    )
    ws.add_data_validation(dv_no)
    dv_no.add(NO_CELL)

    # --- ② または 生徒名で選択 ---
    ws.cell(row=5, column=1, value="② または生徒名で選択 ▶").font = Font(bold=True)
    name_cell = ws.cell(row=5, column=3, value=None)
    name_cell.font = Font(bold=True, size=14)
    name_cell.fill = input_fill
    name_cell.border = input_border
    ws.merge_cells("C5:E5")

    dv_name = DataValidation(
        type="list",
        formula1=f"={name_range_ref}",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="入力エラー",
        error="一覧表にある生徒名から選択してください。",
    )
    ws.add_data_validation(dv_name)
    dv_name.add(NAME_INPUT_CELL)

    # --- 選択中の生徒（①②どちらの入力からでも解決） ---
    ws.cell(row=6, column=1, value="選択中の生徒").font = Font(bold=True)
    resolved_cell = ws.cell(
        row=6,
        column=3,
        value=(
            f'=IF({NO_CELL_ABS}<>"",'
            f'IFERROR(INDEX({TABLE_NAME}[生徒名],MATCH({NO_CELL_ABS},{TABLE_NAME}[No],0)),'
            f'"⚠ 該当するNoがありません"),'
            f'IF({NAME_INPUT_ABS}<>"",{NAME_INPUT_ABS},""))'
        ),
    )
    resolved_cell.font = Font(bold=True, size=13, color="FF223757")
    ws.merge_cells(start_row=6, start_column=3, end_row=6, end_column=5)

    # --- 明細 ---
    detail_rows = [
        ("教室", lookup("教室"), None),
        ("学年", lookup("学年"), None),
        ("コース", lookup("コース"), None),
        ("1コマ単価", lookup("1コマ単価"), MONEY_FORMAT),
        ("週コマ数", lookup("週コマ数"), "0"),
        ("コマ数（月間）", lookup("コマ数"), "0"),
        ("教材費", lookup("教材費"), MONEY_FORMAT),
        ("授業料", lookup("授業料"), MONEY_FORMAT),
        (
            "消費税（教材費の10%）",
            f'IFERROR(ROUND(INDEX({TABLE_NAME}[教材費],MATCH({RESOLVED_CELL_ABS},{TABLE_NAME}[生徒名],0))*0.1,0),"")',
            MONEY_FORMAT,
        ),
        ("ご請求額（税込）", lookup("請求額(税込)"), MONEY_FORMAT),
    ]

    start_row = 9
    ws.cell(row=start_row - 1, column=1, value="ご請求内容").font = Font(bold=True, size=12)
    for i, (label, formula, number_format) in enumerate(detail_rows):
        r = start_row + i
        label_cell = ws.cell(row=r, column=1, value=label)
        label_cell.font = Font(bold=(label == "ご請求額（税込）"))
        value_cell = ws.cell(row=r, column=3, value=f"={formula}")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        if number_format:
            value_cell.number_format = number_format
        if label == "ご請求額（税込）":
            value_cell.font = Font(bold=True, size=14, color="FFC00000")
            label_cell.font = Font(bold=True, size=12)
            for col in range(1, 6):
                ws.cell(row=r, column=col).fill = PatternFill(
                    start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid"
                )

    total_row_ref = start_row + len(detail_rows) - 1  # ご請求額の行

    # --- メール本文（件名・本文の元になる文面。件名行にも本文コピー行にも使い回す） ---
    subject_formula = f'"【{JUKU_NAME}】" & {RESOLVED_CELL_ABS} & "様 月末請求のお知らせ"'
    body_formula = (
        f'IF({RESOLVED_CELL_ABS}="","生徒名を選択すると本文が生成されます。",'
        f'{RESOLVED_CELL_ABS}&"様"&CHAR(10)&CHAR(10)&'
        f'"いつもお世話になっております。{JUKU_NAME}です。"&CHAR(10)&'
        f'"今月分の請求内容は以下の通りです。"&CHAR(10)&CHAR(10)&'
        f'"教室：　"&{lookup("教室")}&CHAR(10)&'
        f'"コース：　"&{lookup("コース")}&CHAR(10)&'
        f'"コマ数：　"&{lookup("コマ数")}&"コマ"&CHAR(10)&'
        f'"授業料：　"&TEXT({lookup("授業料")},"#,##0")&"円"&CHAR(10)&'
        f'"教材費：　"&TEXT({lookup("教材費")},"#,##0")&"円"&CHAR(10)&'
        f'"消費税：　"&TEXT(ROUND({lookup("教材費")}*0.1,0),"#,##0")&"円"&CHAR(10)&'
        f'"----------------"&CHAR(10)&'
        f'"ご請求額：　"&TEXT({lookup("請求額(税込)")},"#,##0")&"円"&CHAR(10)&CHAR(10)&'
        f'"お手数ですが、期日までにお振込みくださいますようお願いいたします。")'
    )

    # --- メール送信 ---
    # ①ワンクリックの mailto リンク（既定のメールアプリが設定済みの場合に動く）
    # ②宛先・件名・本文をそれぞれ単独セルでも表示（コピペで送る場合のフォールバック）
    email_row = total_row_ref + 3
    ws.cell(row=email_row - 1, column=1, value="メールで送る").font = Font(bold=True, size=12)

    ws.cell(row=email_row, column=1, value="宛先（保護者メール）")
    ws.cell(row=email_row, column=3, value=f"={lookup('保護者メール')}")
    ws.merge_cells(start_row=email_row, start_column=3, end_row=email_row, end_column=5)

    subject_row = email_row + 1
    ws.cell(row=subject_row, column=1, value="件名")
    ws.cell(row=subject_row, column=3, value=f"={subject_formula}")
    ws.merge_cells(start_row=subject_row, start_column=3, end_row=subject_row, end_column=5)

    # Gmail の作成画面を直接開く専用URL（mailto: と違い、OS/ブラウザ側の
    # メールアプリ設定が一切不要 — 普通の https リンクとしてブラウザで開く）
    gmail_url_formula = (
        f'"https://mail.google.com/mail/?view=cm&fs=1&to=" & '
        f'ENCODEURL({lookup("保護者メール")}) & "&su=" & ENCODEURL({subject_formula}) & '
        f'"&body=" & ENCODEURL({body_formula})'
    )
    gmail_link_formula = (
        f'IF({RESOLVED_CELL_ABS}="","",'
        f'HYPERLINK({gmail_url_formula},"📧 Gmailでワンクリック作成"))'
    )
    link_row = subject_row + 1
    link_cell = ws.cell(row=link_row, column=3, value=f"={gmail_link_formula}")
    link_cell.font = Font(bold=True, size=12, color="FF1155CC", underline="single")
    link_cell.fill = PatternFill(start_color="FFE2EFDA", end_color="FFE2EFDA", fill_type="solid")
    ws.merge_cells(start_row=link_row, start_column=3, end_row=link_row, end_column=5)

    # クリックが効かない環境向けに、同じURLをそのままコピーできるプレーンテキストでも表示する
    url_row = link_row + 1
    ws.cell(row=url_row, column=1, value="↑クリックで開かない場合はこのURLをコピー")
    ws.cell(row=url_row, column=1).font = Font(size=9, color="FF5B6472")
    url_cell = ws.cell(
        row=url_row,
        column=3,
        value=f'=IF({RESOLVED_CELL_ABS}="","",{gmail_url_formula})',
    )
    url_cell.font = Font(size=9, color="FF5B6472")
    url_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=url_row, start_column=3, end_row=url_row + 1, end_column=6)
    ws.row_dimensions[url_row].height = 30

    note_row = url_row + 2
    ws.cell(
        row=note_row,
        column=3,
        value=(
            "※Googleにログイン済みの状態でブラウザが開く必要があります。"
            "それでも開けない場合は、上の「宛先」「件名」と下の「本文プレビュー」を"
            "コピーして普段お使いのメールに貼り付けてください。"
        ),
    ).font = Font(size=9, color="FF5B6472")
    ws.cell(row=note_row, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=note_row, start_column=3, end_row=note_row, end_column=6)

    # --- 本文プレビュー（コピーもできる） ---
    body_row = note_row + 2
    ws.cell(row=body_row, column=1, value="本文プレビュー（コピーも可）").font = Font(bold=True)

    body_cell = ws.cell(row=body_row + 1, column=1, value=f"={body_formula}")
    ws.merge_cells(start_row=body_row + 1, start_column=1, end_row=body_row + 8, end_column=6)
    body_cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    body_cell.border = BORDER

    # --- 列幅 ---
    for col, width in zip("ABCDEFG", (18, 4, 14, 14, 14, 12, 12)):
        ws.column_dimensions[col].width = width

    ws.sheet_view.showGridLines = False


def main():
    students = build_students()
    wb = build_workbook(students)
    output_path = "一覧表.xlsx"
    wb.save(output_path)
    print(f"生徒{len(students)}名分の月末請求一覧表を作成しました: {output_path}")


if __name__ == "__main__":
    main()
