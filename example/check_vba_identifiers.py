"""ShiftDemo_utf8.txt の簡易点検：未宣言の識別子と、Public Sub の数を調べる。

Excel を動かせない環境で、Option Explicit 違反（宣言忘れ・タイプミス）を見つけるための補助。
各 Sub/Function の中で使われている名前が「ローカル変数・引数・モジュール変数・Const・
プロシージャ名・VBA/Excel の組み込み名」のどれかに当たるかを照合する。
`.` の後ろ（オブジェクトのメンバー）と `名前:=`（名前付き引数）は対象外。

使い方:  python check_vba_identifiers.py
"""
import re
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent / "ShiftDemo_utf8.txt"

# このモジュールで使う VBA / Excel の組み込み名（大文字小文字は区別しない）
BUILTINS = """
attribute vb_name option explicit public private const as string long boolean double variant date object
sub function end exit dim redim erase if then else elseif for to step next each in do loop while wend
and or not mod is nothing true false byval byref optional with set call goto on error resume
trim cstr clng cdbl cdate val int day month year weekday dateserial isdate isnumeric timer format
array mid replace split lbound ubound rnd randomize msgbox debug print err raise description choose
vbobjecterror vbinformation vbexclamation vblf vbsunday vbmonday vbsaturday vbwhite
application thisworkbook worksheet range formatcondition union
xlnone xlcolorindexautomatic xlvalidatelist xlvalidalertstop xlcellvalue xlequal xlexpression
xlcalculationmanual errhandler
""".split()


def strip_code(line):
    """文字列リテラルとコメントを取り除く。"""
    out, in_str = [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 1
                else:
                    in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append(' "" ')
        elif ch == "'":
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(text):
    buf = ""
    for raw in text.splitlines():
        code = strip_code(raw).rstrip()
        if code.endswith(" _"):
            buf += code[:-2] + " "
            continue
        yield buf + code
        buf = ""


DECL_RE = re.compile(r'^\s*(?:Dim|Private|Public|Const|Public Const|Private Const|ReDim)\s+(.*)$', re.I)
PROC_RE = re.compile(r'^\s*(?:Public|Private)?\s*(Sub|Function)\s+(\w+)\s*\((.*?)\)', re.I)


def declared_names(rest):
    """`a As Long, b() As String, c(1 To 3) As Long` から名前を取り出す。"""
    rest = re.sub(r'\([^()]*\)', '', rest)               # 配列の添字を除く
    rest = re.sub(r'=.*$', '', rest) if rest.strip().lower().startswith(("const",)) else rest
    names = []
    for part in rest.split(","):
        part = re.sub(r'^\s*Const\s+', '', part, flags=re.I)
        m = re.match(r'\s*(\w+)', part)
        if m:
            names.append(m.group(1).lower())
    return names


def main():
    text = SRC.read_text(encoding="utf-8")
    lines = list(logical_lines(text))
    module, procs, current = set(), {}, None
    publics = []
    for ln in lines:
        pm = PROC_RE.match(ln)
        if pm:
            current = pm.group(2)
            procs[current] = {"locals": set(), "body": []}
            module.add(current.lower())
            if ln.strip().lower().startswith("public"):
                publics.append(current)
            for p in pm.group(3).split(","):
                m = re.match(r'\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)', p, re.I)
                if m:
                    procs[current]["locals"].add(m.group(1).lower())
            continue
        if re.match(r'^\s*End\s+(Sub|Function)', ln, re.I):
            current = None
            continue
        dm = DECL_RE.match(ln)
        if dm and not re.match(r'^\s*ReDim', ln, re.I):
            target = module if current is None else procs[current]["locals"]
            target.update(declared_names(dm.group(1)))
            if current is None:
                continue
            # 宣言行の Const の右辺などは本文として調べる
        if current is not None:
            procs[current]["body"].append(ln)

    known = set(BUILTINS) | module
    problems = []
    for name, info in procs.items():
        ok = known | info["locals"] | {name.lower()}
        for ln in info["body"]:
            if DECL_RE.match(ln) and not re.match(r'^\s*ReDim', ln, re.I):
                ln = re.sub(r'\bAs\s+\w+', '', ln)
                ln = re.sub(r'^\s*Dim\s+', '', ln)
            ln = re.sub(r'\b\w+:=', '', ln)                     # 名前付き引数
            ln = re.sub(r'\.\s*\w+', '', ln)                    # メンバー
            for ident in re.findall(r'\b[A-Za-z_]\w*\b', ln):
                if ident.lower() not in ok:
                    problems.append(f"{name}: {ident}  ← {ln.strip()[:70]}")

    print(f"プロシージャ {len(procs)} 個、モジュールレベルの名前 {len(module)} 個を照合")
    print(f"Public Sub/Function: {', '.join(publics)}（{len(publics)} 個）")
    if problems:
        print("未宣言の可能性がある識別子:")
        print("\n".join(sorted(set(problems))))
        return 1
    print("未宣言の識別子: なし")
    return 0


if __name__ == "__main__":
    sys.exit(main())
