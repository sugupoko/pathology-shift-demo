"""週休・業務割当表デモ：VBA（ShiftDemo）と同じ手順を Python で再現し、受け入れ基準 1～6 を検証する。

使い方:
    python simulate_shift.py              # シード1の表＋シード1～20の検証 → sim_result.txt
    python simulate_shift.py --seeds 100  # シード1～100で検証
    python simulate_shift.py --xlsx 別のファイル.xlsx

- 行・列の位置やシート名は ShiftDemo_utf8.txt の `Public Const` を読み取って使う。
- 初期データは shift_demo_base.xlsx から読む（Python 側に書き写さない）。
- Python と VBA では乱数列が違うので、配置そのものは Excel の結果と一致しない。
  確かめたいのは「どのシードでも基準を満たすか」。
"""
import argparse
import calendar
import datetime as dt
import random
import re
import sys
import time
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
CONST_RE = re.compile(r'^\s*Public\s+Const\s+(\w+)\s+As\s+(\w+)\s*=\s*(.+)$')


def read_consts(path):
    """`Public Const 名前 As 型 = 値` を辞書にする（文字列と整数）。"""
    consts = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        m = CONST_RE.match(line)
        if not m:
            continue
        name, _typ, rest = m.groups()
        if rest.startswith('"'):
            consts[name] = re.match(r'"((?:[^"]|"")*)"', rest).group(1).replace('""', '"')
        else:
            consts[name] = int(rest.split("'")[0].strip())
    return consts


C = read_consts(HERE / "ShiftDemo_utf8.txt")


def text(v):
    """セルの値を VBA の Trim$(CStr(...)) と同じように文字列にする。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


# ==================================================================
# 読み込み（VBA の LoadAll / ReadSchedule / ApplyRequests に対応）
# ==================================================================
class Model:
    def __init__(self, xlsx):
        wb = load_workbook(xlsx)                     # 数式ではなく値（記号）を読む
        self.load_settings(wb[C["SH_SET"]])
        self.load_kinds(wb[C["SH_KIND"]])
        self.load_staff(wb[C["SH_STAFF"]])
        self.load_calendar(wb[C["SH_SCHED"]], wb[C["SH_SET"]])
        self.requests = self.read_requests(wb[C["SH_REQ"]])
        self.sheet_cells = self.read_schedule(wb[C["SH_SCHED"]])

    def load_settings(self, ws):
        num = lambda r: int(float(text(ws.cell(r, C["SET_COL_VALUE"]).value) or 0))
        self.off_max = num(C["SET_ROW_OFFMAX"]) or C["STAFF_MAX"]
        self.trials = max(1, num(C["SET_ROW_TRIALS"]))
        self.seed = text(ws.cell(C["SET_ROW_SEED"], C["SET_COL_VALUE"]).value)

    def load_kinds(self, ws):
        self.kinds = {}                               # 記号 → (週休に数える, 休み人数に数える, 自動配置)
        self.off_sym = ""
        self.tasks, self.need = [], {}                # need[記号][曜日 0=月..5=土]
        for r in range(C["KIND_ROW1"], C["KIND_ROW1"] + C["KIND_MAX"]):
            s = text(ws.cell(r, C["KC_SYM"]).value)
            if not s:
                continue
            auto = text(ws.cell(r, C["KC_AUTO"]).value)
            self.kinds[s] = (text(ws.cell(r, C["KC_WEEKOFF"]).value) == C["WORD_YES"],
                             text(ws.cell(r, C["KC_OFFCNT"]).value) == C["WORD_YES"], auto)
            if auto == C["WORD_WEEKOFF"] and not self.off_sym:
                self.off_sym = s
            if auto == C["WORD_TASK"] and len(self.tasks) < C["TASK_MAX"]:
                self.tasks.append(s)
                self.need[s] = [int(float(text(ws.cell(r, C["KC_NEED1"] + k).value) or 0)) for k in range(6)]

    def load_staff(self, ws):
        self.names, self.skills, self.satfix, self.can = [], [], [], []
        for r in range(C["STAFF_ROW1"], C["STAFF_ROW1"] + C["STAFF_MAX"]):
            name = text(ws.cell(r, C["SC_NAME"]).value)
            if not name:
                continue
            skill = text(ws.cell(r, C["SC_SKILL"]).value)
            parts = {p.strip() for p in skill.replace("、", ",").replace("，", ",").split(",")}
            self.names.append(name)
            self.skills.append(skill)
            self.satfix.append(text(ws.cell(r, C["SC_TYPE"]).value) == C["WORD_SATFIX"])
            self.can.append({t: (t in parts) for t in self.tasks})

    def load_calendar(self, ws, ws_set):
        self.year = int(ws.cell(C["ROW_YM"], C["COL_YEAR"]).value)
        self.month = int(ws.cell(C["ROW_YM"], C["COL_MONTH"]).value)
        self.days = calendar.monthrange(self.year, self.month)[1]
        hol = {to_date(ws_set.cell(C["HOL_ROW1"] + k, C["HOL_COL"]).value) for k in range(C["HOL_MAX"])}
        self.wday, self.is_work, self.prev_work = {}, {}, {}
        last = 0
        for d in range(1, self.days + 1):
            date = dt.date(self.year, self.month, d)
            self.wday[d] = date.weekday()            # 0=月 … 5=土, 6=日
            self.is_work[d] = date.weekday() != 6 and date not in hol
            self.prev_work[d] = last                 # 日曜・祝日を飛ばした1つ前の稼働日
            if self.is_work[d]:
                last = d
        # 対象週：祝日でない土曜 d について max(1, d-5)～d
        self.weeks = [(max(1, d - 5), d) for d in range(1, self.days + 1)
                      if self.wday[d] == 5 and self.is_work[d]]

    def read_requests(self, ws):
        reqs = []                                    # (職員番号, 日, 記号)
        for r in range(C["REQ_ROW1"], C["REQ_ROW1"] + C["REQ_MAX"]):
            sym = text(ws.cell(r, C["RC_SYM"]).value)
            date = to_date(ws.cell(r, C["RC_DATE"]).value)
            name = text(ws.cell(r, C["RC_NAME"]).value)
            if sym and date and name in self.names and (date.year, date.month) == (self.year, self.month):
                reqs.append((self.names.index(name), date.day, sym))
        return reqs

    def read_schedule(self, ws):
        return [[text(ws.cell(C["ROW_STAFF1"] + i, C["COL_DAY1"] + d).value) for d in range(C["DAY_MAX"])]
                for i in range(C["STAFF_MAX"])]

    # ---- 記号の性質 ----
    def is_off(self, s):
        return s in self.kinds and self.kinds[s][1]

    def is_weekoff(self, s):
        return s in self.kinds and self.kinds[s][0]

    def is_auto(self, s):
        return s in self.kinds and self.kinds[s][2] in (C["WORD_WEEKOFF"], C["WORD_TASK"])

    def need_of(self, t, d):
        return self.need[t][self.wday[d]] if self.wday[d] <= 5 else 0

    # ---- 土台（VBA: ReadSchedule → RemoveAutoSymbols → ApplyRequests）。plan[i][d] は d=1..31 ----
    def base_plan(self):
        base = [[""] + [("" if self.is_auto(s) else s) for s in row] for row in self.sheet_cells]
        for i, d, sym in self.requests:
            base[i][d] = sym
        return base


# ==================================================================
# 週休・業務・評価（VBA の PlaceWeeklyOff / PlaceTasks / CalcPenalty に対応）
# ==================================================================
def place_weekly_off(m, plan, rnd):
    if not m.off_sym:
        return
    for frm, sat in m.weeks:
        for i in range(len(m.names)):
            if any(m.is_weekoff(plan[i][d]) for d in range(frm, sat + 1)):
                continue                             # ◎がある週は追加しない
            if m.satfix[i] and plan[i][sat] == "":
                day = sat                            # 土曜固定の人は土曜
            else:
                day, k = 0, 0                        # 空いている稼働日から1/k方式でランダム
                for d in range(frm, sat + 1):
                    if m.is_work[d] and plan[i][d] == "":
                        k += 1
                        if int(rnd.random() * k) == 0:
                            day = d
            if day > 0:
                plan[i][day] = m.off_sym


def pick_task_staff(m, plan, times, t, d, avoid_repeat, rnd):
    best, fewest, k = -1, 10**9, 0
    p = m.prev_work[d]
    for i in range(len(m.names)):
        ok = m.can[i][t] and plan[i][d] == ""
        if ok and avoid_repeat and p > 0:
            ok = plan[i][p] != t
        if ok:
            if times[i][t] < fewest:
                best, fewest, k = i, times[i][t], 1
            elif times[i][t] == fewest:
                k += 1
                if int(rnd.random() * k) == 0:
                    best = i
    return best


def place_tasks(m, plan, rnd):
    times = [{t: 0 for t in m.tasks} for _ in m.names]
    for d in range(1, m.days + 1):
        if not m.is_work[d]:
            continue
        for t in m.tasks:
            for _ in range(m.need_of(t, d)):
                i = pick_task_staff(m, plan, times, t, d, True, rnd)
                if i < 0:
                    i = pick_task_staff(m, plan, times, t, d, False, rnd)
                if i >= 0:
                    plan[i][d] = t
                    times[i][t] += 1


def task_spread(m, plan, t):
    counts = [sum(plan[i][d] == t for d in range(1, m.days + 1)) for i in range(len(m.names)) if m.can[i][t]]
    return max(counts) - min(counts) if counts else 0


def evaluate(m, plan):
    """ペナルティと、不足（日, 業務, 人数）・休み人数（日→人数）を返す。"""
    shorts, off_num = [], {}
    short_sum = over_sum = repeat_sum = 0
    for d in range(1, m.days + 1):
        off_num[d] = sum(m.is_off(plan[i][d]) for i in range(len(m.names)))
        if not m.is_work[d]:
            continue
        over_sum += max(0, off_num[d] - m.off_max)
        p = m.prev_work[d]
        for t in m.tasks:
            who = [i for i in range(len(m.names)) if plan[i][d] == t]
            if p > 0:
                repeat_sum += sum(plan[i][p] == t for i in who)
            if len(who) < m.need_of(t, d):
                shorts.append((d, t, m.need_of(t, d) - len(who)))
                short_sum += m.need_of(t, d) - len(who)
    spread_sum = sum(task_spread(m, plan, t) for t in m.tasks)
    pen = short_sum * 100 + over_sum * 10 + repeat_sum * 5 + spread_sum
    return pen, shorts, off_num


def update_schedule(m, seed):
    """［更新］と同じ：試行回数だけ繰り返して、ペナルティ最小（同点は先勝ち）の案を返す。"""
    rnd = random.Random(seed)
    base = m.base_plan()
    best, best_pen, tried = None, None, 0
    for trial in range(1, m.trials + 1):
        tried = trial
        plan = [row[:] for row in base]
        place_weekly_off(m, plan, rnd)
        place_tasks(m, plan, rnd)
        pen = evaluate(m, plan)[0]
        if best_pen is None or pen < best_pen:
            best, best_pen = plan, pen
        if pen == 0:
            break
    return best, best_pen, tried


# ==================================================================
# 受け入れ基準（VBA の SelfCheck と同じ判定＋基準6）
# ==================================================================
def check_all(m, plan):
    pen, shorts, off_num = evaluate(m, plan)
    lab = lambda d: f"{m.month}/{d}"
    res = {}
    # 1. 全稼働日に業務が必要人数
    res[1] = [f"{lab(d)} {t} 不足{n}" for d, t, n in shorts]
    # 2. 同じ業務の2稼働日連続なし
    ng = []
    for d in range(1, m.days + 1):
        p = m.prev_work[d]
        if m.is_work[d] and p > 0:
            for i, name in enumerate(m.names):
                if plan[i][d] in m.tasks and plan[i][d] == plan[i][p]:
                    ng.append(f"{name} {plan[i][d]} {lab(p)}と{lab(d)}")
    res[2] = ng
    # 3. 業務回数の差が2以下
    spreads = {t: task_spread(m, plan, t) for t in m.tasks}
    res[3] = [f"{t} の差が {s}" for t, s in spreads.items() if s > C["SPREAD_OK"]]
    # 4. 週休：全員が対象週ごとにちょうど1回（◎が2つの週は除く）、土曜固定の○は土曜
    ng = []
    for frm, sat in m.weeks:
        for i, name in enumerate(m.names):
            syms = [plan[i][d] for d in range(frm, sat + 1) if m.is_weekoff(plan[i][d])]
            n_req = sum(s != m.off_sym for s in syms)
            if len(syms) != 1 and n_req < 2:
                ng.append(f"{name} {lab(frm)}～{lab(sat)} 週休{len(syms)}回")
            if m.satfix[i]:
                for d in range(frm, sat):
                    if plan[i][d] == m.off_sym and plan[i][sat] == "":
                        ng.append(f"{name} {lab(d)} 土曜以外の{m.off_sym}")
    res[4] = ng
    # 5. 希望休が上書きされていない
    res[5] = [f"{m.names[i]} {lab(d)} {s}" for i, d, s in m.requests if plan[i][d] != s]
    # 6. 赤になるのは「希望休だけで上限を超える日」の休み超過だけ（初期データでは 11/20）
    base = m.base_plan()
    expected = {("休み", d) for d in range(1, m.days + 1)
                if m.is_work[d] and sum(m.is_off(base[i][d]) for i in range(len(m.names))) > m.off_max}
    reds = {("業務" + t, d) for d, t, _n in shorts} | \
           {("休み", d) for d in range(1, m.days + 1) if m.is_work[d] and off_num[d] > m.off_max}
    res[6] = [f"{lab(d)} {kind}" for kind, d in sorted(reds ^ expected, key=lambda x: x[1])]
    return res, spreads, reds, pen


TITLES = {1: "全稼働日に 包・仕・切 が各1名", 2: "同じ業務の2稼働日連続なし",
          3: "業務回数の差（最大－最小）が2以下", 4: "週休の回数と土曜固定",
          5: "希望休が上書きされていない", 6: "赤は 11/20 の休み上限超過だけ"}


# ==================================================================
# 表示
# ==================================================================
def width(s):
    return sum(2 if ord(ch) > 0x7F else 1 for ch in s)


def pad(s, w):
    return s + " " * (w - width(s))


def render_table(m, plan):
    wd = "月火水木金土日"
    out = []
    head1 = pad("氏名", 7) + pad("担当", 7) + "".join(f"{d:>4}" for d in range(1, m.days + 1)) + "  休 " + "".join(pad(t, 3) for t in m.tasks)
    head2 = " " * 14 + "".join("  " + wd[m.wday[d]] for d in range(1, m.days + 1))
    out += [head1, head2]
    for i, name in enumerate(m.names):
        cells = ""
        for d in range(1, m.days + 1):
            s = plan[i][d] or ("■" if not m.is_work[d] else "・")
            cells += " " * (4 - width(s)) + s
        offs = sum(m.is_off(plan[i][d]) for d in range(1, m.days + 1))
        cnts = "".join(f"{sum(plan[i][d] == t for d in range(1, m.days + 1)):>3}" for t in m.tasks)
        out.append(pad(name, 7) + pad(m.skills[i] or "－", 7) + cells + f"  {offs:>2}" + cnts)
    _pen, shorts, off_num = evaluate(m, plan)
    for t in m.tasks:
        row = "".join(f"{sum(plan[i][d] == t for i in range(len(m.names))):>4}" if m.is_work[d] else "   ■"
                      for d in range(1, m.days + 1))
        out.append(pad(f"人数 {t}", 14) + row)
    row = "".join((f"{off_num[d]:>3}" + ("!" if off_num[d] > m.off_max else " ")) if m.is_work[d] else "   ■"
                  for d in range(1, m.days + 1))
    out.append(pad("休み人数", 14) + row)
    out.append("（■=日曜・祝日、・=通常勤務、! =休み上限超過）")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=str(HERE / "shift_demo_base.xlsx"))
    ap.add_argument("--seeds", type=int, default=20, help="検証するシードの数（1～N）")
    ap.add_argument("--out", default=str(HERE / "sim_result.txt"))
    args = ap.parse_args()

    m = Model(args.xlsx)
    lines = [f"=== 週休・業務割当表 シミュレーション（{m.year}年{m.month}月、職員{len(m.names)}人、"
             f"試行回数{m.trials}、休み上限{m.off_max}）===",
             f"稼働日 {sum(m.is_work.values())} 日／対象週 " +
             "、".join(f"{m.month}/{a}～{m.month}/{b}" for a, b in m.weeks), ""]

    # シード1（設定シートのシード）の表
    seed0 = int(float(m.seed)) if m.seed else 1
    t0 = time.perf_counter()
    plan, pen, tried = update_schedule(m, seed0)
    sec = time.perf_counter() - t0
    res, spreads, reds, _ = check_all(m, plan)
    lines.append(f"--- シード{seed0} の結果（試行 {tried} 回、ペナルティ {pen}、赤 {len(reds)} 件、{sec:.2f} 秒）---")
    lines += render_table(m, plan)
    lines.append("")
    for k in range(1, 7):
        extra = "（差: " + " ".join(f"{t}={s}" for t, s in spreads.items()) + "）" if k == 3 else ""
        lines.append(f"基準{k} {TITLES[k]}{extra}: {'OK' if not res[k] else 'NG ' + '／'.join(res[k][:5])}")
    lines.append(f"赤のセル: " + "、".join(f"{m.month}/{d} {kind}" for kind, d in sorted(reds, key=lambda x: x[1])))
    lines.append("")

    # シードを変えて検証
    lines.append(f"--- シード1～{args.seeds} での検証 ---")
    lines.append("シード  試行  ペナルティ  赤  秒     基準1 2 3 4 5 6  業務回数の差")
    all_ok, max_sec = True, 0.0
    for seed in range(1, args.seeds + 1):
        t0 = time.perf_counter()
        plan, pen, tried = update_schedule(m, seed)
        sec = time.perf_counter() - t0
        max_sec = max(max_sec, sec)
        res, spreads, reds, _ = check_all(m, plan)
        marks = " ".join("○" if not res[k] else "×" for k in range(1, 7))
        ok = all(not res[k] for k in range(1, 7))
        all_ok &= ok
        lines.append(f"{seed:>5}  {tried:>4}  {pen:>9}  {len(reds):>2}  {sec:5.2f}  {marks}  "
                     + " ".join(f"{t}={s}" for t, s in spreads.items())
                     + ("" if ok else "  NG: " + " / ".join(f"基準{k}:{res[k][0]}" for k in res if res[k])))
    lines.append("")
    lines.append(f"まとめ: {args.seeds} 個のシードで基準1～6 が{'すべて合格' if all_ok else '不合格あり'}"
                 f"（1回あたり最大 {max_sec:.2f} 秒、Python）")

    text_out = "\n".join(lines) + "\n"
    Path(args.out).write_text(text_out, encoding="utf-8")
    sys.stdout.write(text_out)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
