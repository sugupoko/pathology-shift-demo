"""週休・業務割当表デモ：shift_demo_base.xlsx と ShiftDemo.bas を作る。

使い方:  python build_shift_demo.py
  - 行・列の位置やシート名は ShiftDemo_utf8.txt の `Public Const` を読み取って使う
    （Python 側に値を書き写さないので、VBA とレイアウトが食い違わない）。
  - ShiftDemo_utf8.txt（正本）から ShiftDemo.bas（CP932・CRLF）も作る。
データはすべて架空。
"""
import datetime as dt
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
VBA_UTF8 = HERE / "ShiftDemo_utf8.txt"
VBA_BAS = HERE / "ShiftDemo.bas"
XLSX = HERE / "shift_demo_base.xlsx"

# ------------------------------------------------------------------
# VBA の Public Const を読み取る
# ------------------------------------------------------------------
CONST_RE = re.compile(r'^\s*Public\s+Const\s+(\w+)\s+As\s+(\w+)\s*=\s*(.+)$')


def read_consts(path=VBA_UTF8):
    """`Public Const 名前 As 型 = 値` を辞書にする（文字列と整数）。"""
    consts = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        m = CONST_RE.match(line)
        if not m:
            continue
        name, typ, rest = m.groups()
        if rest.startswith('"'):
            s = re.match(r'"((?:[^"]|"")*)"', rest).group(1)
            consts[name] = s.replace('""', '"')
        else:
            consts[name] = int(rest.split("'")[0].strip())
    return consts


C = read_consts()


def vba_color_to_hex(v):
    """VBA の色（BGR の Long）を openpyxl の ARGB 文字列にする。"""
    r, g, b = v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF
    return f"FF{r:02X}{g:02X}{b:02X}"


# ------------------------------------------------------------------
# 初期データ（架空）
# ------------------------------------------------------------------
YEAR, MONTH = 2026, 11
HOLIDAYS = [(dt.date(2026, 11, 3), "文化の日"), (dt.date(2026, 11, 23), "勤労感謝の日")]
SETTINGS = {"OFFMAX": 4, "TRIALS": 400, "SEED": 1}

# 氏名, 担当可能な業務, 週休区分, 備考
STAFF = [
    ("技師A", "包,仕", "通常", "兼務"),
    ("技師B", "包", "通常", ""),
    ("技師C", "包,切", "通常", "兼務"),
    ("技師D", "包", "通常", ""),
    ("技師E", "仕", "通常", ""),
    ("技師F", "仕,切", "通常", "兼務"),
    ("技師G", "仕", "通常", ""),
    ("技師H", "仕", "通常", ""),
    ("技師I", "切", "通常", ""),
    ("技師J", "切", "通常", ""),
    ("技師K", "切", "土曜固定", ""),
    ("技師L", "", "土曜固定", "業務の担当なし"),
]

# 氏名, 日, 記号（すべて 2026年11月）
REQUESTS = [
    ("技師A", 10, "◎"), ("技師B", 5, "◎"), ("技師B", 20, "◎"), ("技師C", 26, "有休"),
    ("技師D", 20, "◎"), ("技師E", 12, "◎"), ("技師E", 20, "◎"), ("技師F", 27, "◎"),
    ("技師G", 20, "◎"), ("技師H", 4, "◎"), ("技師H", 25, "◎"), ("技師I", 13, "◎"),
    ("技師J", 18, "◎"), ("技師L", 20, "◎"), ("技師L", 30, "◎"),
]

YES, NO = C["WORD_YES"], "×"
SIX = [1, 1, 1, 1, 1, 1]            # 月～土の必要人数
# 記号, 名称, 週休に数える, 休み人数に数える, 自動配置, 月～土の必要人数, 文字色, 塗り
KINDS = [
    ("○", "指定休", YES, YES, C["WORD_WEEKOFF"], None, "006100", "C6EFCE"),
    ("◎", "希望休", YES, YES, C["WORD_NONE"], None, "1F4E79", "BDD7EE"),
    ("有休", "有給休暇", NO, YES, C["WORD_NONE"], None, "7030A0", "E4DFEC"),
    ("包", "包埋", NO, NO, C["WORD_TASK"], SIX, "843C0C", "FCE4D6"),
    ("仕", "仕分け", NO, NO, C["WORD_TASK"], SIX, "7F6000", "FFF2CC"),
    ("切", "切り出し", NO, NO, C["WORD_TASK"], SIX, "1E5D5D", "D0F0EE"),
]
TASKS = [k[0] for k in KINDS if k[4] == C["WORD_TASK"]][: C["TASK_MAX"]]
OFF_SYMS = {k[0] for k in KINDS if k[3] == YES}

# ------------------------------------------------------------------
# 書式の部品
# ------------------------------------------------------------------
FONT_NAME = "Meiryo UI"
YELLOW = PatternFill("solid", start_color="FFFFFF99", end_color="FFFFFF99")
HEAD_FILL = PatternFill("solid", start_color="FFDDEBF7", end_color="FFDDEBF7")
THIN = Side(style="thin", color="FFA6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")


def font(**kw):
    return Font(name=FONT_NAME, **kw)


def style_header_row(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.font = font(bold=True)
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = CENTER


def L(col):
    return get_column_letter(col)


# ------------------------------------------------------------------
# 各シート
# ------------------------------------------------------------------
def build_usage(ws):
    lines = [
        ("週休・業務割当表デモ　使い方（データはすべて架空）", True),
        ("", False),
        ("① 「希望休マスタ」に、提出された希望休（氏名・日付・記号）を貼る", False),
        ("② 「勤務表」の［更新］ボタンを押す → 週休（○）と業務（包・仕・切）が自動で入る", False),
        ("③ 赤いセルを確認する（業務の不足、休み人数の上限超過）。必要ならプルダウンで手直し", False),
        ("④ 月を変えるときは、勤務表の年・月（黄色）を変えて［新しい月］を押す", False),
        ("⑤ ［自己チェック］で、ルールを守れているかを確認できる", False),
        ("", False),
        ("※ ボタンが無いときは：VBE で ShiftDemo.bas をインポートし、マクロ「SetupButtons」を1回実行するとボタンが出ます。", False),
        ("※ 手入力した ○ や業務記号は、自動で入れたものと区別できないため［更新］で入れ直されます（◎・有休は残ります）。", False),
        ("※ 手でプルダウンを直すと色はすぐ変わりますが、集計の数字は次の［更新］で変わります。", False),
        ("※ 記号の色を変えるときは、区分マスタの L列（文字の色）・M列（塗りつぶし）の書式を変えて［色・プルダウン反映］。", False),
        ("※ 乱数シード（設定シート）を入れておくと、何度押しても同じ結果になります（空欄なら毎回ランダム）。", False),
    ]
    for r, (text, bold) in enumerate(lines, start=1):
        ws.cell(r, 1, text).font = font(bold=bold, size=14 if bold else 11)
    ws.column_dimensions["A"].width = 120


def build_staff(ws):
    heads = ["氏名", "担当可能な業務", "週休区分", "備考"]
    for c, h in enumerate(heads, start=1):
        ws.cell(1, c, h)
    style_header_row(ws, 1, len(heads))
    for r, row in enumerate(STAFF, start=C["STAFF_ROW1"]):
        vals = {C["SC_NAME"]: row[0], C["SC_SKILL"]: row[1], C["SC_TYPE"]: row[2], C["SC_NOTE"]: row[3]}
        for c, v in vals.items():
            cell = ws.cell(r, c, v if v != "" else None)
            cell.font = font()
            cell.border = BOX
    last = C["STAFF_ROW1"] + C["STAFF_MAX"] - 1
    dv = DataValidation(type="list", formula1=f'"{C["WORD_NORMAL"]},{C["WORD_SATFIX"]}"', allow_blank=True)
    dv.add(f"{L(C['SC_TYPE'])}{C['STAFF_ROW1']}:{L(C['SC_TYPE'])}{last}")
    ws.add_data_validation(dv)
    ws.cell(last + 2, 1, "担当可能な業務はカンマ区切り（例：包,仕）。職員は最大12人。").font = font(color="FF595959")
    for col, w in zip("ABCD", (12, 18, 12, 20)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def build_kinds(ws):
    heads = ["記号", "名称", "週休に数える", "休み人数に数える", "自動配置",
             "月", "火", "水", "木", "金", "土", "文字色", "塗りつぶし色"]
    for c, h in enumerate(heads, start=1):
        ws.cell(1, c, h)
    style_header_row(ws, 1, len(heads))
    for r, (sym, name, wk, off, auto, need, fg, bg) in enumerate(KINDS, start=C["KIND_ROW1"]):
        vals = {C["KC_SYM"]: sym, C["KC_NAME"]: name, C["KC_WEEKOFF"]: wk,
                C["KC_OFFCNT"]: off, C["KC_AUTO"]: auto}
        for k in range(6):
            vals[C["KC_NEED1"] + k] = need[k] if need else None
        for c, v in vals.items():
            cell = ws.cell(r, c, v)
            cell.font = font()
            cell.border = BOX
            cell.alignment = CENTER
        fcell = ws.cell(r, C["KC_FONT"], sym)              # 文字の色を読み取るセル
        fcell.font = font(bold=True, color="FF" + fg)
        fcell.alignment = CENTER
        fcell.border = BOX
        bcell = ws.cell(r, C["KC_FILL"])                    # 塗りつぶしの色を読み取るセル
        bcell.fill = PatternFill("solid", start_color="FF" + bg, end_color="FF" + bg)
        bcell.border = BOX
    last = C["KIND_ROW1"] + C["KIND_MAX"] - 1
    for col in (C["KC_WEEKOFF"], C["KC_OFFCNT"]):
        dv = DataValidation(type="list", formula1=f'"{YES},{NO}"', allow_blank=True)
        dv.add(f"{L(col)}{C['KIND_ROW1']}:{L(col)}{last}")
        ws.add_data_validation(dv)
    dv = DataValidation(type="list", allow_blank=True,
                        formula1=f'"{C["WORD_WEEKOFF"]},{C["WORD_TASK"]},{C["WORD_NONE"]}"')
    dv.add(f"{L(C['KC_AUTO'])}{C['KIND_ROW1']}:{L(C['KC_AUTO'])}{last}")
    ws.add_data_validation(dv)
    note_row = C["KIND_ROW1"] + len(KINDS) + 1
    notes = ["色は、L列のセルの「文字の色」と M列のセルの「塗りつぶしの色」を読み取ります。",
             "色を変えたら勤務表の［色・プルダウン反映］を押してください。赤は警告用なので記号には使わないでください。",
             "自動配置：週休＝［更新］で週1回入れ直す／業務＝［更新］で毎日入れ直す／なし＝入れ直さない。業務は3つまで。"]
    for k, n in enumerate(notes):
        ws.cell(note_row + k, 1, n).font = font(color="FF595959")
    widths = [7, 11, 13, 16, 10] + [5] * 6 + [8, 12]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = "A2"


def build_requests(ws):
    heads = ["氏名", "日付", "記号"]
    for c, h in enumerate(heads, start=1):
        ws.cell(1, c, h)
    style_header_row(ws, 1, len(heads))
    for r, (name, day, sym) in enumerate(REQUESTS, start=C["REQ_ROW1"]):
        ws.cell(r, C["RC_NAME"], name).font = font()
        dcell = ws.cell(r, C["RC_DATE"], dt.date(YEAR, MONTH, day))
        dcell.number_format = "yyyy/m/d"
        dcell.font = font()
        ws.cell(r, C["RC_SYM"], sym).font = font()
        ws.cell(r, C["RC_SYM"]).alignment = CENTER
    ws.cell(1, 5, f"1行に1件。対象年月の行だけ反映されます（最大{C['REQ_MAX']}行）。").font = font(color="FF595959")
    for col, w in zip("ABC", (12, 12, 7)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def build_settings(ws):
    ws.cell(1, 1, "項目")
    ws.cell(1, 2, "値")
    ws.cell(1, C["HOL_COL"], "祝日（会社休日）")
    ws.cell(1, C["HOL_COL"] + 1, "名称")
    for c in (1, 2, C["HOL_COL"], C["HOL_COL"] + 1):
        cell = ws.cell(1, c)
        cell.font = font(bold=True)
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = CENTER
    items = [(C["SET_ROW_OFFMAX"], "1日の休み人数の上限", SETTINGS["OFFMAX"]),
             (C["SET_ROW_TRIALS"], "試行回数", SETTINGS["TRIALS"]),
             (C["SET_ROW_SEED"], "乱数シード（空欄なら毎回ランダム）", SETTINGS["SEED"])]
    for r, label, val in items:
        ws.cell(r, 1, label).font = font()
        v = ws.cell(r, C["SET_COL_VALUE"], val)
        v.font = font(bold=True)
        v.fill = YELLOW
        v.border = BOX
        v.alignment = CENTER
    for k in range(C["HOL_MAX"]):
        r = C["HOL_ROW1"] + k
        dcell = ws.cell(r, C["HOL_COL"])
        dcell.number_format = "yyyy/m/d"
        dcell.border = BOX
        ws.cell(r, C["HOL_COL"] + 1).border = BOX
        if k < len(HOLIDAYS):
            dcell.value = HOLIDAYS[k][0]
            dcell.font = font()
            ws.cell(r, C["HOL_COL"] + 1, HOLIDAYS[k][1]).font = font()
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions[L(C["HOL_COL"])].width = 17
    ws.column_dimensions[L(C["HOL_COL"] + 1)].width = 16


def initial_plan():
    """［新しい月］を押したときと同じ：希望休だけが入った表（職員×日）。"""
    names = [s[0] for s in STAFF][: C["STAFF_MAX"]]
    plan = [["" for _ in range(C["DAY_MAX"])] for _ in names]
    for name, day, sym in REQUESTS:
        if name in names:
            plan[names.index(name)][day - 1] = sym
    return plan


def is_workday(day):
    d = dt.date(YEAR, MONTH, day)
    return d.weekday() != 6 and d not in {h[0] for h in HOLIDAYS}


def build_schedule(ws):
    last_day_col = C["COL_DAY1"] + C["DAY_MAX"] - 1
    staff_last = C["ROW_STAFF1"] + C["STAFF_MAX"] - 1
    yref = f"${L(C['COL_YEAR'])}${C['ROW_YM']}"
    mref = f"${L(C['COL_MONTH'])}${C['ROW_YM']}"

    # タイトル・年月
    ws.cell(C["ROW_TITLE"], 1, "週休・業務割当表（デモ）").font = font(bold=True, size=14)
    ws.row_dimensions[C["ROW_TITLE"]].height = 34
    ws.cell(C["ROW_TITLE"], 1).alignment = Alignment(vertical="center")
    for col, label, val in ((C["COL_YEAR"], "年", YEAR), (C["COL_MONTH"], "月", MONTH)):
        ws.cell(C["ROW_YM"], col - 1, label).font = font(bold=True)
        ws.cell(C["ROW_YM"], col - 1).alignment = Alignment(horizontal="right")
        v = ws.cell(C["ROW_YM"], col, val)
        v.fill = YELLOW
        v.font = font(bold=True, size=12)
        v.border = BOX
        v.alignment = CENTER
    ws.cell(C["ROW_YM"], C["COL_BTN"], "※データはすべて架空です。ボタンはマクロ SetupButtons を1回実行すると出ます。"
            ).font = font(size=9, color="FF7F7F7F")

    # 見出し（氏名・担当可能業務）
    for col, text in ((C["COL_NAME"], "氏名"), (C["COL_SKILL"], "担当可能業務")):
        ws.cell(C["ROW_DATE"], col, text)
        ws.merge_cells(start_row=C["ROW_DATE"], start_column=col, end_row=C["ROW_WDAY"], end_column=col)
        cell = ws.cell(C["ROW_DATE"], col)
        cell.font = font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = CENTER
        for r in (C["ROW_DATE"], C["ROW_WDAY"]):
            ws.cell(r, col).border = BOX

    # 日付行・曜日行（年月から数式で自動計算）
    for day in range(1, C["DAY_MAX"] + 1):
        col = C["COL_DAY1"] + day - 1
        dref = f"{L(col)}{C['ROW_DATE']}"
        dcell = ws.cell(C["ROW_DATE"], col,
                        f'=IF({day}<=DAY(DATE({yref},{mref}+1,0)),DATE({yref},{mref},{day}),"")')
        dcell.number_format = "d"
        dcell.font = font(bold=True)
        wcell = ws.cell(C["ROW_WDAY"], col,
                        f'=IF({dref}="","",CHOOSE(WEEKDAY({dref}),"日","月","火","水","木","金","土"))')
        wcell.font = font()
        for c in (dcell, wcell):
            c.alignment = CENTER
            c.border = BOX

    # 職員の行
    plan = initial_plan()
    for i in range(C["STAFF_MAX"]):
        r = C["ROW_STAFF1"] + i
        ws.row_dimensions[r].height = 20
        name, skill = (STAFF[i][0], STAFF[i][1]) if i < len(STAFF) else ("", "")
        ws.cell(r, C["COL_NAME"], name or None).font = font()
        ws.cell(r, C["COL_SKILL"], skill or None).font = font(size=10)
        ws.cell(r, C["COL_SKILL"]).alignment = CENTER
        for c in (C["COL_NAME"], C["COL_SKILL"]):
            ws.cell(r, c).border = BOX
        for day in range(1, C["DAY_MAX"] + 1):
            cell = ws.cell(r, C["COL_DAY1"] + day - 1)
            if i < len(plan) and plan[i][day - 1]:
                cell.value = plan[i][day - 1]
            cell.alignment = CENTER
            cell.border = BOX
            cell.font = font(size=10)

    # 集計行（業務ごとの人数・休み人数）：VBA と同じく稼働日だけ値を書く
    ws.cell(C["ROW_TASKCNT1"], C["COL_NAME"], "業務ごとの人数").font = font(bold=True, size=10)
    ws.cell(C["ROW_OFFCNT"], C["COL_NAME"], "休み人数").font = font(bold=True, size=10)
    ws.cell(C["ROW_OFFCNT"], C["COL_SKILL"], f"上限{SETTINGS['OFFMAX']}").font = font(size=9, color="FF7F7F7F")
    sum_rows = [C["ROW_TASKCNT1"] + t for t in range(C["TASK_MAX"])] + [C["ROW_OFFCNT"]]
    for t in range(C["TASK_MAX"]):
        ws.cell(C["ROW_TASKCNT1"] + t, C["COL_SKILL"], TASKS[t] if t < len(TASKS) else None)
    for r in sum_rows:
        for c in (C["COL_NAME"], C["COL_SKILL"]):
            ws.cell(r, c).border = BOX
            ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        for day in range(1, C["DAY_MAX"] + 1):
            cell = ws.cell(r, C["COL_DAY1"] + day - 1)
            cell.border = BOX
            cell.alignment = CENTER
            cell.font = font(size=10)
            if not _in_month(day) or not is_workday(day):
                continue
            col_syms = [row[day - 1] for row in plan]
            if r == C["ROW_OFFCNT"]:
                cell.value = sum(s in OFF_SYMS for s in col_syms)
            else:
                t = r - C["ROW_TASKCNT1"]
                cell.value = sum(s == TASKS[t] for s in col_syms) if t < len(TASKS) else None

    # 集計列（休み回数・業務の回数）
    heads = [("休み", "回数")] + [(TASKS[t] if t < len(TASKS) else "", "回数") for t in range(C["TASK_MAX"])]
    for k, (h1, h2) in enumerate(heads):
        col = C["COL_SUM_OFF"] + k
        for r, h in ((C["ROW_DATE"], h1), (C["ROW_WDAY"], h2)):
            cell = ws.cell(r, col, h)
            cell.font = font(bold=True, size=10)
            cell.fill = HEAD_FILL
            cell.alignment = CENTER
            cell.border = BOX
        for i in range(C["STAFF_MAX"]):
            cell = ws.cell(C["ROW_STAFF1"] + i, col)
            cell.border = BOX
            cell.alignment = CENTER
            cell.font = font(size=10)
            if i < len(plan):
                if k == 0:
                    cell.value = sum(s in OFF_SYMS for s in plan[i][: _days()])
                else:
                    cell.value = sum(s == TASKS[k - 1] for s in plan[i][: _days()])

    # 列幅
    ws.column_dimensions[L(C["COL_NAME"])].width = 15
    ws.column_dimensions[L(C["COL_SKILL"])].width = 12
    for col in range(C["COL_DAY1"], last_day_col + 1):
        ws.column_dimensions[L(col)].width = 4.6
    for col in range(C["COL_SUM_OFF"], C["COL_SUM_TASK1"] + C["TASK_MAX"]):
        ws.column_dimensions[L(col)].width = 5.5
    ws.freeze_panes = ws.cell(C["ROW_STAFF1"], C["COL_DAY1"]).coordinate

    # 入力規則（記号のプルダウン）と条件付き書式（ApplyMaster が同じ内容で作り直す）
    sched = f"{L(C['COL_DAY1'])}{C['ROW_STAFF1']}:{L(last_day_col)}{staff_last}"
    dv = DataValidation(type="list", formula1='"' + ",".join(k[0] for k in KINDS) + '"', allow_blank=True)
    dv.add(sched)
    ws.add_data_validation(dv)
    top = f"{L(C['COL_DAY1'])}${C['ROW_DATE']}"           # C$4
    gray_range = " ".join([
        f"{L(C['COL_DAY1'])}{C['ROW_DATE']}:{L(last_day_col)}{staff_last}",
        f"{L(C['COL_DAY1'])}{C['ROW_TASKCNT1']}:{L(last_day_col)}{C['ROW_TASKCNT1'] + C['TASK_MAX'] - 1}",
        f"{L(C['COL_DAY1'])}{C['ROW_OFFCNT']}:{L(last_day_col)}{C['ROW_OFFCNT']}",
    ])
    gray = vba_color_to_hex(C["COLOR_GRAY"])
    ws.conditional_formatting.add(gray_range, FormulaRule(
        formula=[f'IF({top}="",TRUE,OR(WEEKDAY({top})=1,COUNTIF({C["NAME_HOLIDAYS"]},{top})>0))'],
        fill=PatternFill("solid", start_color=gray, end_color=gray, bgColor=gray), stopIfTrue=True))
    for sym, _n, _w, _o, _a, _need, fg, bg in KINDS:
        ws.conditional_formatting.add(sched, CellIsRule(
            operator="equal", formula=[f'"{sym}"'], font=Font(color="FF" + fg, bold=True),
            fill=PatternFill("solid", start_color="FF" + bg, end_color="FF" + bg, bgColor="FF" + bg)))

    # 印刷設定（横向き・1ページ幅）
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _days():
    return (dt.date(YEAR + (MONTH == 12), MONTH % 12 + 1, 1) - dt.timedelta(days=1)).day


def _in_month(day):
    return day <= _days()


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = C["SH_USAGE"]
    build_usage(ws)
    build_schedule(wb.create_sheet(C["SH_SCHED"]))
    build_staff(wb.create_sheet(C["SH_STAFF"]))
    build_kinds(wb.create_sheet(C["SH_KIND"]))
    build_requests(wb.create_sheet(C["SH_REQ"]))
    build_settings(wb.create_sheet(C["SH_SET"]))
    col = L(C["HOL_COL"])
    ref = f"'{C['SH_SET']}'!${col}${C['HOL_ROW1']}:${col}${C['HOL_ROW1'] + C['HOL_MAX'] - 1}"
    wb.defined_names[C["NAME_HOLIDAYS"]] = DefinedName(C["NAME_HOLIDAYS"], attr_text=ref)
    wb.active = 1                      # 開いたときに勤務表を表示
    for w in wb.worksheets:
        w.sheet_view.tabSelected = (w.title == C["SH_SCHED"])
    wb.calculation.fullCalcOnLoad = True
    wb.save(XLSX)
    print(f"作成: {XLSX.name}")


# ------------------------------------------------------------------
# ShiftDemo.bas（CP932・CRLF）を正本から作る
# ------------------------------------------------------------------
NG_CHARS = {"〜": "〜(U+301C)→～", "−": "−(U+2212)→－", "‖": "‖", "¢": "¢",
            "£": "£", "¬": "¬", "—": "—(U+2014)"}


def make_bas():
    text = VBA_UTF8.read_text(encoding="utf-8")
    for ch, label in NG_CHARS.items():
        if ch in text:
            raise SystemExit(f"CP932 で往復しない文字があります: {label}")
    if not text.startswith('Attribute VB_Name = "ShiftDemo"'):
        raise SystemExit("先頭に Attribute VB_Name がありません")
    lf = text.replace("\r\n", "\n").replace("\r", "\n")
    if not lf.endswith("\n"):
        lf += "\n"
    data = lf.replace("\n", "\r\n").encode("cp932")    # 変換できない文字があればここで止まる
    VBA_BAS.write_bytes(data)
    # 読み戻して確認
    back = VBA_BAS.read_bytes()
    assert back.decode("cp932").replace("\r\n", "\n") == lf, "CP932 で読み戻した内容が一致しない"
    assert back.count(b"\n") == back.count(b"\r\n"), "CRLF でない改行がある"
    assert b"\r" not in back.replace(b"\r\n", b""), "単独の CR がある"
    n_lines = back.count(b"\r\n")
    print(f"作成: {VBA_BAS.name}（CP932・CRLF、{n_lines} 行、読み戻し一致）")


if __name__ == "__main__":
    make_bas()
    build_workbook()
