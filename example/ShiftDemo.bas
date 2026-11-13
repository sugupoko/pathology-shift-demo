Attribute VB_Name = "ShiftDemo"
'==================================================================
' 週休・業務割当表デモ（データはすべて架空）
'   ルールの順に当てはめ、ランダムに何度も試して一番よい案を採用する。
'   ボタン：［更新］［色・プルダウン反映］［新しい月］［クリア］［自己チェック］
'   ※ボタンは、このモジュールを読み込んだあと SetupButtons を1回実行すると出る。
'==================================================================
Option Explicit

'------------------------------------------------------------------
' シート名と、マスタで使う決まった言葉
'------------------------------------------------------------------
Public Const SH_USAGE As String = "使い方"
Public Const SH_SCHED As String = "勤務表"
Public Const SH_STAFF As String = "職員マスタ"
Public Const SH_KIND As String = "区分マスタ"
Public Const SH_REQ As String = "希望休マスタ"
Public Const SH_SET As String = "設定"
Public Const WORD_YES As String = "○"            ' 「週休に数える」「休み人数に数える」の印
Public Const WORD_WEEKOFF As String = "週休"     ' 自動配置：週休
Public Const WORD_TASK As String = "業務"        ' 自動配置：業務
Public Const WORD_NONE As String = "なし"        ' 自動配置：なし
Public Const WORD_SATFIX As String = "土曜固定"  ' 週休区分：土曜固定
Public Const WORD_NORMAL As String = "通常"      ' 週休区分：通常
Public Const NAME_HOLIDAYS As String = "HolidayList"  ' 祝日一覧につける名前

'------------------------------------------------------------------
' 勤務表の位置（行・列の番号）
'------------------------------------------------------------------
Public Const ROW_TITLE As Long = 1       ' タイトルとボタンの行
Public Const ROW_YM As Long = 2          ' 年・月の行
Public Const COL_YEAR As Long = 2        ' 年（B2）
Public Const COL_MONTH As Long = 4       ' 月（D2）
Public Const ROW_DATE As Long = 4        ' 日付の行
Public Const ROW_WDAY As Long = 5        ' 曜日の行
Public Const ROW_STAFF1 As Long = 6      ' 職員の先頭行
Public Const STAFF_MAX As Long = 12      ' 職員の行数（6～17行目）
Public Const ROW_TASKCNT1 As Long = 19   ' 業務ごとの人数の先頭行
Public Const TASK_MAX As Long = 3        ' 業務は3つまで（19～21行目）
Public Const ROW_OFFCNT As Long = 22     ' 休み人数の行
Public Const COL_NAME As Long = 1        ' A列：氏名
Public Const COL_SKILL As Long = 2       ' B列：担当可能業務
Public Const COL_DAY1 As Long = 3        ' C列：1日
Public Const DAY_MAX As Long = 31        ' 31列固定（C～AG列）
Public Const COL_SUM_OFF As Long = 34    ' AH列：休み回数
Public Const COL_SUM_TASK1 As Long = 35  ' AI～AK列：業務の回数
Public Const COL_BTN As Long = 6         ' ボタンを並べ始める列（F列）

'------------------------------------------------------------------
' 区分マスタの位置（1行目は見出し）
'------------------------------------------------------------------
Public Const KIND_ROW1 As Long = 2       ' 先頭行
Public Const KIND_MAX As Long = 30       ' 最大の行数
Public Const KC_SYM As Long = 1          ' A列：記号
Public Const KC_NAME As Long = 2         ' B列：名称
Public Const KC_WEEKOFF As Long = 3      ' C列：週休に数える
Public Const KC_OFFCNT As Long = 4       ' D列：休み人数に数える
Public Const KC_AUTO As Long = 5         ' E列：自動配置
Public Const KC_NEED1 As Long = 6        ' F～K列：月～土の必要人数
Public Const KC_FONT As Long = 12        ' L列：このセルの文字の色を使う
Public Const KC_FILL As Long = 13        ' M列：このセルの塗りつぶしの色を使う

'------------------------------------------------------------------
' 職員マスタ・希望休マスタ・設定の位置
'------------------------------------------------------------------
Public Const STAFF_ROW1 As Long = 2      ' 職員マスタの先頭行（2～13行目）
Public Const SC_NAME As Long = 1         ' A列：氏名
Public Const SC_SKILL As Long = 2        ' B列：担当可能な業務
Public Const SC_TYPE As Long = 3         ' C列：週休区分
Public Const SC_NOTE As Long = 4         ' D列：備考
Public Const REQ_ROW1 As Long = 2        ' 希望休マスタの先頭行
Public Const REQ_MAX As Long = 500       ' 希望休マスタの最大の行数
Public Const RC_NAME As Long = 1         ' A列：氏名
Public Const RC_DATE As Long = 2         ' B列：日付
Public Const RC_SYM As Long = 3          ' C列：記号
Public Const SET_COL_VALUE As Long = 2   ' 設定の値はB列
Public Const SET_ROW_OFFMAX As Long = 2  ' B2：1日の休み人数の上限
Public Const SET_ROW_TRIALS As Long = 3  ' B3：試行回数
Public Const SET_ROW_SEED As Long = 4    ' B4：乱数シード（空欄なら毎回ランダム）
Public Const HOL_ROW1 As Long = 2        ' 祝日一覧の先頭行（D2）
Public Const HOL_MAX As Long = 40        ' 祝日一覧の行数（D2～D41）
Public Const HOL_COL As Long = 4         ' D列：祝日（E列：名称）

'------------------------------------------------------------------
' 色と判定の基準
'------------------------------------------------------------------
Public Const COLOR_GRAY As Long = 12566463   ' 灰色 RGB(191,191,191)
Public Const COLOR_RED As Long = 3947775     ' 警告の赤 RGB(255,60,60)
Public Const SPREAD_OK As Long = 2           ' 業務回数の差（最大－最小）の合格ライン
Public Const LIST_MAX As Long = 5            ' 自己チェックで表示するNG箇所の数

'==================================================================
' 作業用の変数（Sub どうしで受け渡す）
'==================================================================
' カレンダー
Private mYear As Long, mMonth As Long, mDays As Long
Private mWday(1 To DAY_MAX) As Long           ' 曜日（1=日 ～ 7=土）
Private mIsWork(1 To DAY_MAX) As Boolean      ' 稼働日か（日曜・祝日以外）
Private mPrevWork(1 To DAY_MAX) As Long       ' 1つ前の稼働日（無ければ0）
Private mWeekCount As Long                    ' 対象週の数
Private mWeekFrom(1 To 6) As Long             ' 対象週の初日
Private mWeekSat(1 To 6) As Long              ' 対象週の土曜
' 区分マスタ
Private mKindCount As Long
Private mKindRow(1 To KIND_MAX) As Long       ' 区分マスタの何行目か
Private mKindSym(1 To KIND_MAX) As String     ' 記号
Private mKindWeekOff(1 To KIND_MAX) As Boolean ' 週休に数える
Private mKindOff(1 To KIND_MAX) As Boolean    ' 休み人数に数える
Private mKindAuto(1 To KIND_MAX) As String    ' 週休／業務／なし
Private mOffSym As String                     ' 自動で入れる週休の記号（○）
Private mTaskCount As Long
Private mTaskSym(1 To TASK_MAX) As String     ' 業務の記号（包・仕・切）
Private mTaskNeed(1 To TASK_MAX, 1 To 7) As Long  ' 曜日別の必要人数
' 職員マスタ
Private mStaffCount As Long
Private mStaffName(1 To STAFF_MAX) As String
Private mStaffSkill(1 To STAFF_MAX) As String
Private mSatFix(1 To STAFF_MAX) As Boolean    ' 土曜固定か
Private mCanDo(1 To STAFF_MAX, 1 To TASK_MAX) As Boolean  ' 業務を担当できるか
' 設定
Private mOffMax As Long, mTrials As Long
Private mSeed As Variant
' 勤務表（職員×日の記号）
Private mBase() As String                     ' 希望休などを入れた土台
Private mPlan() As String                     ' 試している案
Private mBest() As String                     ' いちばんよい案
Private mTaskTimes(1 To STAFF_MAX, 1 To TASK_MAX) As Long  ' 試行中の業務回数
' 評価の結果
Private mShort(1 To TASK_MAX, 1 To DAY_MAX) As Long  ' 業務の不足人数
Private mOffNum(1 To DAY_MAX) As Long                ' 休み人数
' 画面更新などの元の設定
Private mFastOn As Boolean, mOldCalc As Long, mOldEvents As Boolean

'==================================================================
' ボタンから呼ぶマクロ（6つ）
'==================================================================

' ［更新］：週休と業務を自動で割り当て、条件を満たせない日を赤くする
Public Sub UpdateSchedule()
    Dim startTime As Double, trial As Long, tried As Long
    Dim pen As Long, bestPen As Long, redCount As Long
    On Error GoTo ErrHandler
    startTime = Timer
    BeginFast

    ' 1. 読み込み
    LoadAll
    ReadSchedule
    RemoveAutoSymbols          ' ○と業務を消す（◎・有休は残す）
    ApplyRequests              ' 希望休マスタを反映
    InitRandom

    ' 2～4. 週休 → 業務 → 評価 を試行回数だけ繰り返す
    bestPen = -1
    For trial = 1 To mTrials
        tried = trial
        mPlan = mBase
        PlaceWeeklyOff
        PlaceTasks
        pen = CalcPenalty()
        If bestPen < 0 Or pen < bestPen Then
            bestPen = pen
            mBest = mPlan
        End If
        If pen = 0 Then Exit For
    Next

    ' 5. 書き出し
    mPlan = mBest
    CalcPenalty                ' 採用案の不足人数・休み人数を数え直す
    WriteSchedule
    WriteSummary
    ClearRed
    redCount = MarkRed()
    EndFast

    MsgBox "更新しました。" & vbLf & _
        "試行回数：" & tried & " 回" & vbLf & _
        "ペナルティ：" & bestPen & vbLf & _
        "赤のセル：" & redCount & " 件" & vbLf & _
        "所要時間：" & Format(Timer - startTime, "0.00") & " 秒", vbInformation, "更新"
    Exit Sub
ErrHandler:
    ShowError
End Sub

' ［色・プルダウン反映］：区分マスタの記号と色を勤務表に反映する
Public Sub ApplyMaster()
    On Error GoTo ErrHandler
    BeginFast
    ApplyMasterCore
    EndFast
    MsgBox "区分マスタの記号と色を反映しました。", vbInformation, "色・プルダウン反映"
    Exit Sub
ErrHandler:
    ShowError
End Sub

' ［新しい月］：記号をすべて消し、新しい年月の希望休を反映する（赤は付けない）
Public Sub CreateNewMonth()
    On Error GoTo ErrHandler
    BeginFast
    LoadAll
    ReDim mBase(1 To STAFF_MAX, 1 To DAY_MAX)   ' 空の表から始める
    ApplyRequests
    mPlan = mBase
    CalcPenalty
    WriteSchedule
    WriteSummary
    ClearRed
    EndFast
    MsgBox mYear & "年" & mMonth & "月の表にしました（希望休を反映済み）。" & vbLf & _
        "［更新］を押すと週休と業務が入ります。", vbInformation, "新しい月"
    Exit Sub
ErrHandler:
    ShowError
End Sub

' ［クリア］：自動配置した記号（○と業務）だけを消す（録画の撮り直し用）
Public Sub ClearSchedule()
    On Error GoTo ErrHandler
    BeginFast
    LoadAll
    ReadSchedule
    RemoveAutoSymbols
    mPlan = mBase
    CalcPenalty
    WriteSchedule
    WriteSummary
    ClearRed
    EndFast
    MsgBox "自動配置した記号を消しました。", vbInformation, "クリア"
    Exit Sub
ErrHandler:
    ShowError
End Sub

' 勤務表に5つのボタンを並べてマクロを割り当てる（再実行しても重複しない）
Public Sub SetupButtons()
    Dim ws As Worksheet, macros As Variant, captions As Variant, widths As Variant
    Dim k As Long, s As Long, x As Double, btn As Object
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    macros = Array("UpdateSchedule", "ApplyMaster", "CreateNewMonth", "ClearSchedule", "SelfCheck")
    captions = Array("更新", "色・プルダウン反映", "新しい月", "クリア", "自己チェック")
    widths = Array(70, 130, 80, 60, 95)

    ' 前に作った btn_マクロ名 のボタンだけを消す
    For k = 0 To UBound(macros)
        For s = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(s).Name = "btn_" & macros(k) Then ws.Shapes(s).Delete
        Next
    Next

    ' F列の左端から右へ1行目に並べる
    x = ws.Cells(ROW_TITLE, COL_BTN).Left
    For k = 0 To UBound(macros)
        Set btn = ws.Buttons.Add(x, ws.Cells(ROW_TITLE, 1).Top + 4, widths(k), ws.Rows(ROW_TITLE).Height - 8)
        btn.Name = "btn_" & macros(k)
        btn.OnAction = macros(k)
        btn.Caption = captions(k)
        btn.Font.Size = 11
        x = x + widths(k) + 6
    Next

    BeginFast
    ApplyMasterCore
    EndFast
    MsgBox "ボタンを配置しました。", vbInformation, "ボタンの配置"
    Exit Sub
ErrHandler:
    ShowError
End Sub

' ［自己チェック］：受け入れ基準1～5を、勤務表の今の内容で判定して表示する
Public Sub SelfCheck()
    Dim msg As String
    On Error GoTo ErrHandler
    BeginFast
    LoadAll
    ReadSchedule
    mPlan = mBase
    CalcPenalty                 ' 不足人数と休み人数を数える
    msg = CheckTasksFilled() & CheckNoRepeat() & CheckSpread() & _
          CheckWeeklyOff() & CheckRequests() & ListOverDays()
    EndFast
    Debug.Print msg
    MsgBox msg, vbInformation, "自己チェック（" & mYear & "年" & mMonth & "月）"
    Exit Sub
ErrHandler:
    ShowError
End Sub

'==================================================================
' 読み込み
'==================================================================

' マスタ・設定・カレンダーをまとめて読み込む
Private Sub LoadAll()
    LoadSettings
    LoadKinds
    LoadStaff
    LoadCalendar
End Sub

' 設定シートの値を読み込む
Private Sub LoadSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SET)
    mOffMax = CLng(Val(CStr(ws.Cells(SET_ROW_OFFMAX, SET_COL_VALUE).Value)))
    If mOffMax <= 0 Then mOffMax = STAFF_MAX
    mTrials = CLng(Val(CStr(ws.Cells(SET_ROW_TRIALS, SET_COL_VALUE).Value)))
    If mTrials < 1 Then mTrials = 1
    mSeed = ws.Cells(SET_ROW_SEED, SET_COL_VALUE).Value
End Sub

' 区分マスタを配列にまとめて読み込む
Private Sub LoadKinds()
    Dim v As Variant, r As Long, wd As Long, s As String
    v = ThisWorkbook.Worksheets(SH_KIND).Cells(KIND_ROW1, 1).Resize(KIND_MAX, KC_FILL).Value
    mKindCount = 0: mTaskCount = 0: mOffSym = ""
    Erase mTaskSym, mTaskNeed
    For r = 1 To KIND_MAX
        s = Trim$(CStr(v(r, KC_SYM)))
        If s <> "" Then
            mKindCount = mKindCount + 1
            mKindRow(mKindCount) = KIND_ROW1 + r - 1
            mKindSym(mKindCount) = s
            mKindWeekOff(mKindCount) = (Trim$(CStr(v(r, KC_WEEKOFF))) = WORD_YES)
            mKindOff(mKindCount) = (Trim$(CStr(v(r, KC_OFFCNT))) = WORD_YES)
            mKindAuto(mKindCount) = Trim$(CStr(v(r, KC_AUTO)))
            If mKindAuto(mKindCount) = WORD_WEEKOFF And mOffSym = "" Then mOffSym = s
            If mKindAuto(mKindCount) = WORD_TASK And mTaskCount < TASK_MAX Then
                mTaskCount = mTaskCount + 1
                mTaskSym(mTaskCount) = s
                For wd = vbMonday To vbSaturday        ' F列=月 … K列=土
                    mTaskNeed(mTaskCount, wd) = CLng(Val(CStr(v(r, KC_NEED1 + wd - vbMonday))))
                Next
            End If
        End If
    Next
End Sub

' 職員マスタを読み込み、誰がどの業務を担当できるかを調べる
Private Sub LoadStaff()
    Dim v As Variant, r As Long, t As Long, s As String
    v = ThisWorkbook.Worksheets(SH_STAFF).Cells(STAFF_ROW1, 1).Resize(STAFF_MAX, SC_NOTE).Value
    mStaffCount = 0
    Erase mCanDo
    For r = 1 To STAFF_MAX
        s = Trim$(CStr(v(r, SC_NAME)))
        If s <> "" Then
            mStaffCount = mStaffCount + 1
            mStaffName(mStaffCount) = s
            mStaffSkill(mStaffCount) = Trim$(CStr(v(r, SC_SKILL)))
            mSatFix(mStaffCount) = (Trim$(CStr(v(r, SC_TYPE))) = WORD_SATFIX)
            For t = 1 To mTaskCount
                mCanDo(mStaffCount, t) = HasSkill(mStaffSkill(mStaffCount), mTaskSym(t))
            Next
        End If
    Next
End Sub

' 年・月と祝日一覧から、曜日・稼働日・対象週を求める
Private Sub LoadCalendar()
    Dim ws As Worksheet, hol As Variant, d As Long, k As Long, lastWork As Long
    Dim dt As Date, isHoliday As Boolean
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    mYear = CLng(Val(CStr(ws.Cells(ROW_YM, COL_YEAR).Value)))
    mMonth = CLng(Val(CStr(ws.Cells(ROW_YM, COL_MONTH).Value)))
    If mYear < 1900 Or mMonth < 1 Or mMonth > 12 Then
        Err.Raise vbObjectError + 1, , "勤務表の年・月（黄色のセル）を確認してください。"
    End If
    mDays = Day(DateSerial(mYear, mMonth + 1, 0))
    hol = ThisWorkbook.Worksheets(SH_SET).Cells(HOL_ROW1, HOL_COL).Resize(HOL_MAX, 1).Value

    For d = 1 To DAY_MAX
        mWday(d) = 0: mIsWork(d) = False: mPrevWork(d) = 0
        If d <= mDays Then
            dt = DateSerial(mYear, mMonth, d)
            mWday(d) = Weekday(dt)
            isHoliday = False
            For k = 1 To HOL_MAX
                If IsDate(hol(k, 1)) Then
                    If Int(CDate(hol(k, 1))) = dt Then isHoliday = True
                End If
            Next
            mIsWork(d) = (mWday(d) <> vbSunday) And Not isHoliday
            mPrevWork(d) = lastWork                   ' 日曜・祝日を飛ばした1つ前の稼働日
            If mIsWork(d) Then lastWork = d
        End If
    Next

    ' 対象週：祝日でない土曜 d について、max(1, d－5)～d
    mWeekCount = 0
    For d = 1 To mDays
        If mWday(d) = vbSaturday And mIsWork(d) Then
            mWeekCount = mWeekCount + 1
            mWeekSat(mWeekCount) = d
            mWeekFrom(mWeekCount) = d - 5
            If mWeekFrom(mWeekCount) < 1 Then mWeekFrom(mWeekCount) = 1
        End If
    Next
End Sub

' 勤務表の記号を配列 mBase に一括で読み込む
Private Sub ReadSchedule()
    Dim v As Variant, i As Long, d As Long
    v = SchedRange().Value
    ReDim mBase(1 To STAFF_MAX, 1 To DAY_MAX)
    For i = 1 To STAFF_MAX
        For d = 1 To DAY_MAX
            mBase(i, d) = Trim$(CStr(v(i, d)))
        Next
    Next
End Sub

' 「入れ直す」記号（自動配置が週休・業務のもの）を mBase から消す
Private Sub RemoveAutoSymbols()
    Dim i As Long, d As Long, k As Long
    For i = 1 To STAFF_MAX
        For d = 1 To DAY_MAX
            k = KindIndex(mBase(i, d))
            If k > 0 Then
                If mKindAuto(k) = WORD_WEEKOFF Or mKindAuto(k) = WORD_TASK Then mBase(i, d) = ""
            End If
        Next
    Next
End Sub

' 希望休マスタのうち、対象年月の行を mBase に書き込む
Private Sub ApplyRequests()
    Dim v As Variant, r As Long, i As Long, dt As Date
    v = ThisWorkbook.Worksheets(SH_REQ).Cells(REQ_ROW1, 1).Resize(REQ_MAX, RC_SYM).Value
    For r = 1 To REQ_MAX
        If Trim$(CStr(v(r, RC_SYM))) <> "" And IsDate(v(r, RC_DATE)) Then
            dt = CDate(v(r, RC_DATE))
            i = StaffIndex(Trim$(CStr(v(r, RC_NAME))))  ' 職員マスタにない氏名は0
            If i > 0 And Year(dt) = mYear And Month(dt) = mMonth Then
                mBase(i, Day(dt)) = Trim$(CStr(v(r, RC_SYM)))
            End If
        End If
    Next
End Sub

' 乱数を初期化する（シードがあれば毎回同じ結果になる）
Private Sub InitRandom()
    If Trim$(CStr(mSeed)) <> "" And IsNumeric(mSeed) Then
        Rnd -1
        Randomize CDbl(mSeed)
    Else
        Randomize
    End If
End Sub

'==================================================================
' 週休の配置
'==================================================================

' 対象週ごとに、各職員へ週休（○）を1回ずつ入れる
Private Sub PlaceWeeklyOff()
    Dim w As Long, i As Long, d As Long, sat As Long
    If mOffSym = "" Then Exit Sub
    For w = 1 To mWeekCount
        sat = mWeekSat(w)
        For i = 1 To mStaffCount
            If Not HasWeekOff(i, w) Then             ' ◎がある週は追加しない
                If mSatFix(i) And mPlan(i, sat) = "" Then
                    d = sat                          ' 土曜固定の人は土曜
                Else
                    d = PickFreeDay(i, w)            ' 空いている稼働日からランダム
                End If
                If d > 0 Then mPlan(i, d) = mOffSym
            End If
        Next
    Next
End Sub

' 職員 i が対象週 w に「週休に数える」記号をすでに持っているか
Private Function HasWeekOff(ByVal i As Long, ByVal w As Long) As Boolean
    Dim d As Long
    For d = mWeekFrom(w) To mWeekSat(w)
        If IsWeekOffSym(mPlan(i, d)) Then HasWeekOff = True
    Next
End Function

' 対象週 w の空いている稼働日から1日をランダムに選ぶ（無ければ0）
Private Function PickFreeDay(ByVal i As Long, ByVal w As Long) As Long
    Dim d As Long, k As Long
    For d = mWeekFrom(w) To mWeekSat(w)
        If mIsWork(d) And mPlan(i, d) = "" Then
            k = k + 1
            If Int(Rnd * k) = 0 Then PickFreeDay = d   ' k日目を確率1/kで採用
        End If
    Next
End Function

'==================================================================
' 業務の配置
'==================================================================

' 稼働日ごとに、業務（包・仕・切）を必要人数ずつ入れる
Private Sub PlaceTasks()
    Dim d As Long, t As Long, n As Long, i As Long
    Erase mTaskTimes
    For d = 1 To mDays
        If mIsWork(d) Then
            For t = 1 To mTaskCount                       ' 区分マスタの上から
                For n = 1 To mTaskNeed(t, mWday(d))
                    i = PickTaskStaff(t, d, True)         ' まず連続を避けて選ぶ
                    If i = 0 Then i = PickTaskStaff(t, d, False)  ' 誰もいなければ連続も許す
                    If i > 0 Then
                        mPlan(i, d) = mTaskSym(t)
                        mTaskTimes(i, t) = mTaskTimes(i, t) + 1
                    End If
                Next
            Next
        End If
    Next
End Sub

' 業務 t を d 日に入れる人を選ぶ（回数の少ない人を優先、同数ならランダム。無ければ0）
Private Function PickTaskStaff(ByVal t As Long, ByVal d As Long, ByVal avoidRepeat As Boolean) As Long
    Dim i As Long, k As Long, p As Long, fewest As Long, ok As Boolean
    fewest = 999999
    p = mPrevWork(d)
    For i = 1 To mStaffCount
        ok = mCanDo(i, t) And mPlan(i, d) = ""          ' 担当可能で、その日が空欄
        If ok And avoidRepeat And p > 0 Then ok = (mPlan(i, p) <> mTaskSym(t))
        If ok Then
            If mTaskTimes(i, t) < fewest Then
                fewest = mTaskTimes(i, t)
                PickTaskStaff = i
                k = 1
            ElseIf mTaskTimes(i, t) = fewest Then
                k = k + 1
                If Int(Rnd * k) = 0 Then PickTaskStaff = i  ' 同数のk人目を確率1/kで採用
            End If
        End If
    Next
End Function

'==================================================================
' 評価
'==================================================================

' 案 mPlan のペナルティを計算し、不足人数と休み人数も記録する
Private Function CalcPenalty() As Long
    Dim d As Long, t As Long, i As Long, p As Long, cnt As Long
    Dim shortSum As Long, overSum As Long, repeatSum As Long, spreadSum As Long
    Erase mShort, mOffNum
    For d = 1 To mDays
        For i = 1 To mStaffCount
            If IsOffSym(mPlan(i, d)) Then mOffNum(d) = mOffNum(d) + 1
        Next
        If mIsWork(d) Then                                ' 不足と超過は稼働日だけ
            If mOffNum(d) > mOffMax Then overSum = overSum + mOffNum(d) - mOffMax
            p = mPrevWork(d)
            For t = 1 To mTaskCount
                cnt = 0
                For i = 1 To mStaffCount
                    If mPlan(i, d) = mTaskSym(t) Then
                        cnt = cnt + 1
                        If p > 0 Then
                            If mPlan(i, p) = mTaskSym(t) Then repeatSum = repeatSum + 1
                        End If
                    End If
                Next
                If cnt < mTaskNeed(t, mWday(d)) Then
                    mShort(t, d) = mTaskNeed(t, mWday(d)) - cnt
                    shortSum = shortSum + mShort(t, d)
                End If
            Next
        End If
    Next
    For t = 1 To mTaskCount
        spreadSum = spreadSum + TaskSpread(t)
    Next
    CalcPenalty = shortSum * 100 + overSum * 10 + repeatSum * 5 + spreadSum
End Function

' 業務 t について、担当可能な職員の月内回数の「最大－最小」
Private Function TaskSpread(ByVal t As Long) As Long
    Dim i As Long, c As Long, most As Long, least As Long
    most = -1: least = 999999
    For i = 1 To mStaffCount
        If mCanDo(i, t) Then
            c = CountSym(i, mTaskSym(t))
            If c > most Then most = c
            If c < least Then least = c
        End If
    Next
    If most >= 0 Then TaskSpread = most - least
End Function

'==================================================================
' 書き出し
'==================================================================

' 配列 mPlan を勤務表へ一括で書き出す（氏名・担当可能業務も書き直す）
Private Sub WriteSchedule()
    Dim ws As Worksheet, v() As Variant, nameCol() As Variant, skillCol() As Variant
    Dim i As Long, d As Long
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    ReDim v(1 To STAFF_MAX, 1 To DAY_MAX)
    ReDim nameCol(1 To STAFF_MAX, 1 To 1)
    ReDim skillCol(1 To STAFF_MAX, 1 To 1)
    For i = 1 To STAFF_MAX
        For d = 1 To DAY_MAX
            If mPlan(i, d) <> "" Then v(i, d) = mPlan(i, d)
        Next
        If i <= mStaffCount Then
            nameCol(i, 1) = mStaffName(i)
            skillCol(i, 1) = mStaffSkill(i)
        End If
    Next
    SchedRange().Value = v
    ws.Cells(ROW_STAFF1, COL_NAME).Resize(STAFF_MAX, 1).Value = nameCol
    ws.Cells(ROW_STAFF1, COL_SKILL).Resize(STAFF_MAX, 1).Value = skillCol
End Sub

' 集計行（業務ごとの人数・休み人数）と集計列（休み回数・業務の回数）に値を書く
Private Sub WriteSummary()
    Dim ws As Worksheet, taskRow() As Variant, offRow() As Variant, sumCol() As Variant
    Dim i As Long, d As Long, t As Long, c As Long
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    ReDim taskRow(1 To TASK_MAX, 1 To DAY_MAX)
    ReDim offRow(1 To 1, 1 To DAY_MAX)
    ReDim sumCol(1 To STAFF_MAX, 1 To 1 + TASK_MAX)

    ' 日ごとの人数（稼働日だけ書く）
    For d = 1 To mDays
        If mIsWork(d) Then
            offRow(1, d) = mOffNum(d)
            For t = 1 To mTaskCount
                c = 0
                For i = 1 To mStaffCount
                    If mPlan(i, d) = mTaskSym(t) Then c = c + 1
                Next
                taskRow(t, d) = c
            Next
        End If
    Next

    ' 職員ごとの回数
    For i = 1 To mStaffCount
        c = 0
        For d = 1 To mDays
            If IsOffSym(mPlan(i, d)) Then c = c + 1
        Next
        sumCol(i, 1) = c
        For t = 1 To mTaskCount
            sumCol(i, 1 + t) = CountSym(i, mTaskSym(t))
        Next
    Next

    ws.Cells(ROW_TASKCNT1, COL_DAY1).Resize(TASK_MAX, DAY_MAX).Value = taskRow
    ws.Cells(ROW_OFFCNT, COL_DAY1).Resize(1, DAY_MAX).Value = offRow
    ws.Cells(ROW_STAFF1, COL_SUM_OFF).Resize(STAFF_MAX, 1 + TASK_MAX).Value = sumCol
    For t = 1 To TASK_MAX                                ' 業務の記号の見出し
        ws.Cells(ROW_TASKCNT1 + t - 1, COL_SKILL).Value = mTaskSym(t)
        ws.Cells(ROW_DATE, COL_SUM_TASK1 + t - 1).Value = mTaskSym(t)
    Next
End Sub

' 前回の赤を消して元の見た目に戻す
Private Sub ClearRed()
    Dim t As Long
    ResetLook DayRow(ROW_DATE), True                     ' 日付の行だけ太字
    ResetLook DayRow(ROW_OFFCNT), False
    For t = 1 To TASK_MAX
        ResetLook DayRow(ROW_TASKCNT1 + t - 1), False
    Next
End Sub

' 塗りつぶしなし・自動の文字色に戻す
Private Sub ResetLook(ByVal rng As Range, ByVal isBold As Boolean)
    rng.Interior.Pattern = xlNone
    rng.Font.ColorIndex = xlColorIndexAutomatic
    rng.Font.Bold = isBold
End Sub

' 業務が不足した日と、休みが上限を超えた日を赤くする（赤の件数を返す）
Private Function MarkRed() As Long
    Dim ws As Worksheet, d As Long, t As Long, n As Long
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    For d = 1 To mDays
        If mIsWork(d) Then
            For t = 1 To mTaskCount
                If mShort(t, d) > 0 Then
                    PaintRed ws.Cells(ROW_TASKCNT1 + t - 1, COL_DAY1 + d - 1)
                    n = n + 1
                End If
            Next
            If mOffNum(d) > mOffMax Then
                PaintRed ws.Cells(ROW_OFFCNT, COL_DAY1 + d - 1)
                PaintRed ws.Cells(ROW_DATE, COL_DAY1 + d - 1)  ' 日付セルは件数に数えない
                n = n + 1
            End If
        End If
    Next
    MarkRed = n
End Function

' セルを警告の赤（白の太字）にする
Private Sub PaintRed(ByVal cell As Range)
    cell.Interior.Color = COLOR_RED
    cell.Font.Color = vbWhite
    cell.Font.Bold = True
End Sub

'==================================================================
' 区分マスタの反映（プルダウン・色分け・灰色）
'==================================================================

' 区分マスタの記号と色を、勤務表の入力規則と条件付き書式に反映する
Private Sub ApplyMasterCore()
    Dim ws As Worksheet, wsKind As Worksheet, fc As FormatCondition
    Dim k As Long, list As String, topCell As String
    LoadKinds
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    Set wsKind = ThisWorkbook.Worksheets(SH_KIND)

    ' 祝日一覧の名前の定義を作り直す
    ThisWorkbook.Names.Add Name:=NAME_HOLIDAYS, RefersTo:="='" & SH_SET & "'!" & _
        ThisWorkbook.Worksheets(SH_SET).Cells(HOL_ROW1, HOL_COL).Resize(HOL_MAX, 1).Address

    ' 相対参照がずれないよう、勤務表の左上の日付セル（C4）を選んでから作る
    ThisWorkbook.Activate
    ws.Activate
    ws.Cells(ROW_DATE, COL_DAY1).Select

    ' 記号のプルダウン
    For k = 1 To mKindCount
        list = list & "," & mKindSym(k)
    Next
    With SchedRange().Validation
        .Delete
        If mKindCount > 0 Then .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Mid$(list, 2)
    End With

    ' 色分け（記号ごと）
    GrayRange().FormatConditions.Delete
    For k = 1 To mKindCount
        Set fc = SchedRange().FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:="=""" & mKindSym(k) & """")
        fc.Font.Color = wsKind.Cells(mKindRow(k), KC_FONT).Font.Color
        fc.Interior.Color = wsKind.Cells(mKindRow(k), KC_FILL).Interior.Color
    Next

    ' 日曜・祝日・月末を超える列を灰色に（いちばん優先）
    topCell = ws.Cells(ROW_DATE, COL_DAY1).Address(RowAbsolute:=True, ColumnAbsolute:=False)  ' C$4
    Set fc = GrayRange().FormatConditions.Add(Type:=xlExpression, Formula1:= _
        "=IF(" & topCell & "="""",TRUE,OR(WEEKDAY(" & topCell & ")=1,COUNTIF(" & _
        NAME_HOLIDAYS & "," & topCell & ")>0))")
    fc.Interior.Color = COLOR_GRAY
    fc.SetFirstPriority
    fc.StopIfTrue = True
End Sub

'==================================================================
' 自己チェック（受け入れ基準1～5と、参考の休み上限超過）
'==================================================================

' 基準1：全稼働日に業務が必要人数ずつ入っているか
Private Function CheckTasksFilled() As String
    Dim d As Long, t As Long, list As String, cnt As Long
    For d = 1 To mDays
        For t = 1 To mTaskCount
            If mShort(t, d) > 0 Then AddItem list, cnt, DayLabel(d) & " " & mTaskSym(t) & " 不足" & mShort(t, d)
        Next
    Next
    CheckTasksFilled = ResultText("基準1 全稼働日に業務が入る", list, cnt)
End Function

' 基準2：同じ人に同じ業務が2稼働日連続で入っていないか
Private Function CheckNoRepeat() As String
    Dim d As Long, t As Long, i As Long, p As Long, list As String, cnt As Long
    For d = 1 To mDays
        p = mPrevWork(d)
        If mIsWork(d) And p > 0 Then
            For i = 1 To mStaffCount
                For t = 1 To mTaskCount
                    If mPlan(i, d) = mTaskSym(t) And mPlan(i, p) = mTaskSym(t) Then
                        AddItem list, cnt, mStaffName(i) & " " & mTaskSym(t) & " " & DayLabel(p) & "と" & DayLabel(d)
                    End If
                Next
            Next
        End If
    Next
    CheckNoRepeat = ResultText("基準2 同じ業務の連続なし", list, cnt)
End Function

' 基準3：業務ごとの回数の差（最大－最小）が2以下か（差の値も表示）
Private Function CheckSpread() As String
    Dim t As Long, s As Long, values As String, list As String, cnt As Long
    For t = 1 To mTaskCount
        s = TaskSpread(t)
        values = values & " " & mTaskSym(t) & "=" & s
        If s > SPREAD_OK Then AddItem list, cnt, mTaskSym(t) & " の差が " & s
    Next
    CheckSpread = ResultText("基準3 業務回数の差が" & SPREAD_OK & "以下（差:" & values & "）", list, cnt)
End Function

' 基準4：土曜固定の○は土曜に、全員が対象週ごとに週休ちょうど1回か
Private Function CheckWeeklyOff() As String
    Dim w As Long, i As Long, d As Long, nAll As Long, nReq As Long, list As String, cnt As Long
    For w = 1 To mWeekCount
        For i = 1 To mStaffCount
            nAll = 0: nReq = 0
            For d = mWeekFrom(w) To mWeekSat(w)
                If IsWeekOffSym(mPlan(i, d)) Then
                    nAll = nAll + 1
                    If mPlan(i, d) <> mOffSym Then nReq = nReq + 1   ' ◎など
                End If
            Next
            If nAll <> 1 And nReq < 2 Then                 ' ◎が同じ週に2つある場合は除く
                AddItem list, cnt, mStaffName(i) & " " & WeekLabel(w) & " 週休" & nAll & "回"
            End If
            If mSatFix(i) Then                             ' 土曜以外の○は、土曜がふさがっている週だけ認める
                For d = mWeekFrom(w) To mWeekSat(w) - 1
                    If mPlan(i, d) = mOffSym And mPlan(i, mWeekSat(w)) = "" Then
                        AddItem list, cnt, mStaffName(i) & " " & DayLabel(d) & " 土曜以外の" & mOffSym
                    End If
                Next
            End If
        Next
    Next
    CheckWeeklyOff = ResultText("基準4 週休の回数と土曜固定", list, cnt)
End Function

' 基準5：希望休マスタの記号が上書きされずに残っているか
Private Function CheckRequests() As String
    Dim v As Variant, r As Long, i As Long, dt As Date, s As String, list As String, cnt As Long
    v = ThisWorkbook.Worksheets(SH_REQ).Cells(REQ_ROW1, 1).Resize(REQ_MAX, RC_SYM).Value
    For r = 1 To REQ_MAX
        s = Trim$(CStr(v(r, RC_SYM)))
        If s <> "" And IsDate(v(r, RC_DATE)) Then
            dt = CDate(v(r, RC_DATE))
            i = StaffIndex(Trim$(CStr(v(r, RC_NAME))))
            If i > 0 And Year(dt) = mYear And Month(dt) = mMonth Then
                If mPlan(i, Day(dt)) <> s Then AddItem list, cnt, mStaffName(i) & " " & DayLabel(Day(dt)) & " " & s
            End If
        End If
    Next
    CheckRequests = ResultText("基準5 希望休が残っている", list, cnt)
End Function

' （参考）休みが上限を超えた稼働日の一覧
Private Function ListOverDays() As String
    Dim d As Long, s As String
    For d = 1 To mDays
        If mIsWork(d) And mOffNum(d) > mOffMax Then s = s & " " & DayLabel(d) & "（" & mOffNum(d) & "人）"
    Next
    If s = "" Then s = " なし"
    ListOverDays = vbLf & "（参考）休み上限 " & mOffMax & " 人を超えた日：" & s
End Function

' NG の箇所を LIST_MAX 件まで書きためる
Private Sub AddItem(ByRef list As String, ByRef cnt As Long, ByVal item As String)
    cnt = cnt + 1
    If cnt <= LIST_MAX Then list = list & "　・" & item & vbLf
End Sub

' 判定結果を「OK」または「NG（件数）＋箇所」の文章にする
Private Function ResultText(ByVal title As String, ByVal list As String, ByVal cnt As Long) As String
    If cnt = 0 Then
        ResultText = "【OK】" & title & vbLf
    Else
        ResultText = "【NG】" & title & "（" & cnt & "件）" & vbLf & list
        If cnt > LIST_MAX Then ResultText = ResultText & "　・ほか " & (cnt - LIST_MAX) & " 件" & vbLf
    End If
End Function

'==================================================================
' 小さな部品
'==================================================================

' 勤務表の記号を入れる範囲（職員×31日）
Private Function SchedRange() As Range
    Set SchedRange = ThisWorkbook.Worksheets(SH_SCHED).Cells(ROW_STAFF1, COL_DAY1).Resize(STAFF_MAX, DAY_MAX)
End Function

' 勤務表の指定した行の、1日～31日の範囲
Private Function DayRow(ByVal r As Long) As Range
    Set DayRow = ThisWorkbook.Worksheets(SH_SCHED).Cells(r, COL_DAY1).Resize(1, DAY_MAX)
End Function

' 灰色にする範囲（日付行～職員の行と、集計行）
Private Function GrayRange() As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_SCHED)
    Set GrayRange = Union( _
        ws.Cells(ROW_DATE, COL_DAY1).Resize(ROW_STAFF1 + STAFF_MAX - ROW_DATE, DAY_MAX), _
        ws.Cells(ROW_TASKCNT1, COL_DAY1).Resize(TASK_MAX, DAY_MAX), DayRow(ROW_OFFCNT))
End Function

' 記号が区分マスタの何番目か（無ければ0）
Private Function KindIndex(ByVal s As String) As Long
    Dim k As Long
    If s = "" Then Exit Function
    For k = 1 To mKindCount
        If mKindSym(k) = s Then KindIndex = k
    Next
End Function

' 「休み人数に数える」記号か
Private Function IsOffSym(ByVal s As String) As Boolean
    Dim k As Long
    k = KindIndex(s)
    If k > 0 Then IsOffSym = mKindOff(k)
End Function

' 「週休に数える」記号か
Private Function IsWeekOffSym(ByVal s As String) As Boolean
    Dim k As Long
    k = KindIndex(s)
    If k > 0 Then IsWeekOffSym = mKindWeekOff(k)
End Function

' 氏名が職員マスタの何番目か（無ければ0）
Private Function StaffIndex(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To mStaffCount
        If mStaffName(i) = s Then StaffIndex = i
    Next
End Function

' 担当可能業務（カンマ区切り）に記号が含まれるか
Private Function HasSkill(ByVal skill As String, ByVal sym As String) As Boolean
    Dim parts As Variant, k As Long
    skill = Replace(Replace(skill, "、", ","), "，", ",")
    parts = Split(skill, ",")
    For k = LBound(parts) To UBound(parts)
        If Trim$(parts(k)) = sym Then HasSkill = True
    Next
End Function

' 職員 i の月内で、記号 s が入っている日数
Private Function CountSym(ByVal i As Long, ByVal s As String) As Long
    Dim d As Long
    For d = 1 To mDays
        If mPlan(i, d) = s Then CountSym = CountSym + 1
    Next
End Function

' 日付の表示（例：11/20）
Private Function DayLabel(ByVal d As Long) As String
    DayLabel = mMonth & "/" & d
End Function

' 対象週の表示（例：11/16～11/21）
Private Function WeekLabel(ByVal w As Long) As String
    WeekLabel = DayLabel(mWeekFrom(w)) & "～" & DayLabel(mWeekSat(w))
End Function

' 処理を速くするため、画面更新・イベント・自動計算を止める
Private Sub BeginFast()
    If mFastOn Then Exit Sub
    mOldCalc = Application.Calculation
    mOldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    mFastOn = True
End Sub

' 止めていた画面更新・イベント・自動計算を元に戻す
Private Sub EndFast()
    If Not mFastOn Then Exit Sub
    Application.Calculation = mOldCalc
    Application.EnableEvents = mOldEvents
    Application.ScreenUpdating = True
    mFastOn = False
End Sub

' エラーの内容を表示する（設定は必ず元に戻す）
Private Sub ShowError()
    Dim msg As String
    msg = Err.Description
    EndFast
    MsgBox "エラーが起きました：" & vbLf & msg, vbExclamation, "エラー"
End Sub
