Attribute VB_Name = "表格标题_5_预检查"
'======================== 模块：表格标题_5_预检查 ========================
Option Explicit
Private Const BAR_MAX As Integer = 200   ' ProgressForm 进度条满宽

'——用于给窗体读取的孤儿段列表：每行 [1]=标题级别(1..4/0), [2]=文本, [3]=段起点
Public gOrphanRows As Variant

Public Function GetOrphanRows() As Variant
    GetOrphanRows = gOrphanRows
End Function


'（一）入口：自检（生成数组 → 打开窗体）
Public Sub 自检_表格标题样式一致性()
    Dim doc As Document: Set doc = ActiveDocument
    Dim arrTableInfo As Variant
    Dim headStarts() As Long, headKeys() As String, headLvls() As Long, headNums() As String

    ' 1) 打开进度窗（早绑定直接用你的 ProgressForm）
    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "表格预检查 - 初始化…"
    pf.Show vbModeless
    pf.TextBoxStatus.text = "正在开始表格标题预检查"
    DoEvents
    
    
    Dim t As Double: t = t0()                     ' ★ 计时起点
    PF_Log pf, "① 关闭屏幕刷新..."
    Application.ScreenUpdating = False
    PF_Tick pf, t, "关闭屏幕刷新", 10              ' ★ 打点

    ' 2) 建索引（≈30%）
    PF_Log pf, "② 建立标题索引（扫描正文段落）..."
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT
    PF_Tick pf, t, "建立标题索引", 70              ' ★ 打点

    ' 3) 生成表数组（≈70%）
    PF_Log pf, "③ 构建表信息数组（扫描表格与表前段、孤儿段）..."
    arrTableInfo = BuildTableInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT
    PF_Tick pf, t, "构建表信息数组", 160            ' ★ 打点
    
    
    PF_Log pf, "④ 渲染报告到 WebBrowser..."
    Application.ScreenUpdating = True
    PF_Tick pf, t, "打开屏幕刷新", 170

    ' 4) 交给窗体
    Dim frm As 自检报告
    Set frm = New 自检报告
    frm.LoadReportFromArray arrTableInfo
    frm.Show vbModeless
    PF_Tick pf, t, "加载/写入 HTML", 200

    ' 5) 完成
    pf.UpdateProgressBar BAR_MAX, "完成。"
    DoEvents
ABORT:
    'Unload pf
    
    
End Sub
'（一补）入口：自检（图片标题）——逻辑同表格，只是改用图片数组 + 报告切到“图”
Public Sub 自检_图片标题样式一致性()
    Dim doc As Document: Set doc = ActiveDocument
    Dim arrPicInfo As Variant
    Dim headStarts() As Long, headKeys() As String, headLvls() As Long, headNums() As String

    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "图片预检查 - 初始化…"
    pf.Show vbModeless
    pf.TextBoxStatus.text = "正在开始图片标题预检查"
    DoEvents

    Application.ScreenUpdating = False

    ' ① 建索引
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT

    ' ② 生成图片数组
    arrPicInfo = BuildImageInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT

    Application.ScreenUpdating = True

    ' ③ 打开窗体（注意传 "图"）
    Dim frm As 自检报告
    Set frm = New 自检报告
    frm.LoadReportFromArray arrPicInfo, "图"
    frm.Show vbModeless

    pf.UpdateProgressBar BAR_MAX, "完成。"
    DoEvents
ABORT:
    Unload pf
End Sub



'（二）建立“标题索引”（一次扫描，带进度：0%→30%）
Private Sub BuildHeadingIndex(ByVal doc As Document, _
                              ByRef headStarts() As Long, ByRef headKeys() As String, _
                              ByRef headLvls() As Long, ByRef headNums() As String, _
                              ByRef pf As progressForm)

    Dim total As Long: total = doc.Paragraphs.Count
    Dim ts() As Long, tl() As Long, tn() As String, tk() As String
    ReDim ts(1 To total): ReDim tl(1 To total): ReDim tn(1 To total): ReDim tk(1 To total)

    Dim i As Long, cnt As Long
    Dim p As Paragraph

    ' 先给“空索引”最小占位，避免后续未初始化
    ReDim headStarts(1 To 1): headStarts(1) = 0
    ReDim headLvls(1 To 1): headLvls(1) = 0
    ReDim headNums(1 To 1): headNums(1) = ""
    ReDim headKeys(1 To 1): headKeys(1) = ""

    For Each p In doc.Paragraphs
        cnt = cnt + 1

        ' ——抽样刷新进度（每200段）
        If cnt Mod 200 = 0 Then
            Dim w As Integer
            w = CInt(BAR_MAX * 0.3 * cnt / IIf(total = 0, 1, total))
            pf.UpdateProgressBar w, "① 建立标题索引… " & cnt & "/" & total
            DoEvents
            If pf.stopFlag Then Exit Sub
        End If

        ' ——筛选标题1~4
        Dim sty As String: sty = SafeStyleName(p.Range)
        Dim lvl As Long: lvl = HeadingLevelByStyle(sty)
        If lvl >= 1 And lvl <= 4 Then
            i = i + 1
            ts(i) = p.Range.Start
            tl(i) = lvl
            tn(i) = SafeListString(p.Range)
            tk(i) = NormalizeChapterKey(tn(i), lvl)
        End If
    Next

    If i = 0 Then Exit Sub

    ReDim headStarts(1 To i)
    ReDim headLvls(1 To i)
    ReDim headNums(1 To i)
    ReDim headKeys(1 To i)
    Dim k As Long
    For k = 1 To i
        headStarts(k) = ts(k)
        headLvls(k) = tl(k)
        headNums(k) = tn(k)
        headKeys(k) = tk(k)
    Next

    pf.UpdateProgressBar CInt(BAR_MAX * 0.3), "① 建立标题索引完成。"
    DoEvents
End Sub

'（三）生成“表信息数组”（带进度：30%→100%）
'   列定义（1-based）：保持原 1~12 列不变，新增：
'   [13] 本章“孤儿段”HTML（不在任何表前，但样式=【表格标题】）
'   “孤儿段”的归属章按 headStarts/headKeys 最近标题上界判定
'===================================================
'=========================================================
' 构建表格预检查主数组 + 孤儿段清单
' 返回：arr(1..N, 1..12)
'   1=表索引  2=表Range.Start  3=表前段起点  4=表前段文本
'   5=表前段样式  6=是否命中“表格标题”  7=最近标题级别
'   8=最近标题文本  9=最近标题编号串  10=章键  11=章内序号
'   12=表号（如“表3.1-2”）
' 同时输出：gOrphanRows(1..M,1..3) → [lvl, text, start]
'=========================================================
Private Function BuildTableInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    '（零）基本变量准备
    Dim N As Long: N = doc.Tables.Count
    Dim arr As Variant
    Dim basePx As Long: basePx = 70           ' 本阶段进度基线（与入口对应）
    Dim i As Long, tbl As Table
    Dim t0 As Double                          ' 单步计时用
    Dim sumPrev As Double, sumHead As Double, sumWrite As Double ' 分项累计

    '（一）无表格时：仍需收集孤儿段，但主数组返回 Empty
    If N = 0 Then
        BuildTableInfoArray = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "未检测到任何表格（仍将收集孤儿段）。"
    Else
        ReDim arr(1 To N, 1 To 12)
    End If

    '（二）收集所有“样式=表格标题”的段（后续做孤儿段差集用）
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    t0 = Timer
    Call CollectAllCaptionParas(doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt, pf)
    If Not pf Is Nothing Then PF_StepWarn pf, t0, "②-预取所有‘表格标题’段", 1, IIf(N = 0, 1, N), 0.08, 8, basePx, 10

    '（三）逐表生成行：并记录“有效的表前标题段”（字典）
    Dim dictSeq As Object:      Set dictSeq = CreateObject("Scripting.Dictionary")   ' 章键→章内序号
    Dim dictValidCap As Object: Set dictValidCap = CreateObject("Scripting.Dictionary") ' 有效表题起点→True

    For i = 1 To N
        ' ——进度（总进度条：70→150）
        If Not pf Is Nothing Then
            If (i Mod 10 = 0) Or (i = N) Then
                pf.UpdateProgressBar basePx + CInt((80# / IIf(N = 0, 1, N)) * i), "③ 扫描表格… " & i & "/" & N
            End If
        End If

        Set tbl = doc.Tables(i)
        Dim tStart As Long: tStart = tbl.Range.Start

        '（1）就近“上一非空段” → 作为表题段候选
        t0 = Timer
        Dim pStart As Long: pStart = PrevNonEmptyParaStart_ByStart(doc, tStart)
        sumPrev = sumPrev + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "① 找上一非空段", i, N, 0.12, 10, basePx, 80

        '（2）读取该段的文本与样式
        t0 = Timer
        Dim pTxt As String, pSty As String
        If pStart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, pStart))
            pSty = SafeStyleNameByStart(pStart)
        Else
            pTxt = "": pSty = ""
        End If
        sumWrite = sumWrite + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "② 读取表题段文本/样式", i, N, 0.08, 12, basePx, 80

        ' ——命中“表格标题”则记为有效表题（用于后面孤儿段差集）
        If (pStart > 0) And (pSty = "表格标题") Then
            dictValidCap(CStr(pStart)) = True
        End If

        '（3）定位“最近标题”（用预构建的 head* 数组二分上界）
        t0 = Timer
        Dim idx As Long: idx = UpperBoundByStart(headStarts, tStart)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If
        sumHead = sumHead + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "③ 二分定位最近标题", i, N, 0.1, 10, basePx, 80

        '（4）章内序号与“表号”计算（章键按 headKeys）
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "表" & key & "-" & CStr(seq)

        '（5）是否命中【表格标题】样式
        Dim isCap As Boolean: isCap = (pSty = "表格标题")

        '（6）写入主数组（保持原 1~12 列定义不变）
        arr(i, 1) = i
        arr(i, 2) = tStart
        arr(i, 3) = pStart
        arr(i, 4) = pTxt
        arr(i, 5) = pSty
        arr(i, 6) = isCap
        arr(i, 7) = lvl
        arr(i, 8) = hText
        arr(i, 9) = num
        arr(i, 10) = key
        arr(i, 11) = seq
        arr(i, 12) = label
    Next i

    '（四）生成“孤儿段”二维数组（套了表格标题样式，但不在任何表前）
    '     ——采用“按倍数扩容”的缓冲区，避免 O(n^2) 的 ReDim Preserve
    Dim tOrp As Double: tOrp = Timer
    If Not pf Is Nothing Then PF_Log pf, "③-2 收集孤儿段…"

    Dim orCnt As Long: orCnt = 0
    Dim cap As Long, orBuf As Variant
    If capCnt > 0 Then
        cap = IIf(capCnt \ 4 < 8, 8, capCnt \ 4)   ' 初始容量：capCnt 的 1/4，至少 8
        ReDim orBuf(1 To 3, 1 To cap)
        Dim j As Long
        For j = 1 To capCnt
            Dim cs As Long: cs = capStarts(j)
            If cs <= 0 Then GoTo NEXT_J

            If Not dictValidCap.exists(CStr(cs)) Then
                ' ——不在表前：即为孤儿段
                Dim lvl2 As Long, idx2 As Long
                idx2 = UpperBoundByStart(headStarts, cs)
                If idx2 >= 1 Then lvl2 = headLvls(idx2) Else lvl2 = 0

                ' 扩容（2 倍）并写入
                orCnt = orCnt + 1
                If orCnt > cap Then
                    cap = cap * 2
                    ReDim Preserve orBuf(1 To 3, 1 To cap)
                End If
                orBuf(1, orCnt) = lvl2
                orBuf(2, orCnt) = capTexts(j)
                orBuf(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If

    '（五）输出给窗体使用的 gOrphanRows（压缩到 行×3）
    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, k As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For k = 1 To orCnt
            outArr(k, 1) = orBuf(1, k)   ' lvl
            outArr(k, 2) = orBuf(2, k)   ' text
            outArr(k, 3) = orBuf(3, k)   ' start
        Next k
        gOrphanRows = outArr
    End If

    If Not pf Is Nothing Then
        pf.UpdateProgressBar 160, "③-2 孤儿段完成，用时 " & Format(Timer - tOrp, "0.000") & " s"
        pf.UpdateProgressBar 160, "③ 总结：prev=" & Format(sumPrev, "0.000") & "s；head=" & _
                                   Format(sumHead, "0.000") & "s；read/write=" & Format(sumWrite, "0.000") & "s"
    End If

    '（六）返回主数组
    BuildTableInfoArray = arr
End Function






'——（四）工具：收集文档中所有样式=【表格标题】的段落（位置、文本、所属章键）
'Private Sub CollectAllCaptionParas(ByVal doc As Document, _
'                                   'ByRef headStarts() As Long, ByRef headKeys() As String, _
'                                   'ByRef capStarts() As Long, ByRef capTexts() As String, _
'                                   'ByRef capKeys() As String, ByRef capCnt As Long, _
'                                   'Optional ByVal styleName As String = "表格标题")

    'Dim total As Long: total = doc.Paragraphs.Count
    'ReDim capStarts(1 To total)
    'ReDim capTexts(1 To total)
    'ReDim capKeys(1 To total)
    'Dim p As Paragraph, i As Long

    'For Each p In doc.Paragraphs
        'If SafeStyleName(p.Range) = styleName Then
            'i = i + 1
            'capStarts(i) = p.Range.Start
            'capTexts(i) = TrimVisible(FirstParaTextAtStart(doc, p.Range.Start))
            'Dim idx As Long: idx = UpperBoundByStart(headStarts, p.Range.Start)
            'If idx >= 1 Then
                'capKeys(i) = headKeys(idx)
            'Else
                'capKeys(i) = "0"
            'End If
        'End If
    'Next
    'If i = 0 Then
        'ReDim capStarts(1 To 1): capStarts(1) = 0
        'ReDim capTexts(1 To 1):  capTexts(1) = ""
        'ReDim capKeys(1 To 1):   capKeys(1) = "0"
        'capCnt = 0
    'Else
        'ReDim Preserve capStarts(1 To i)
        'ReDim Preserve capTexts(1 To i)
        'ReDim Preserve capKeys(1 To i)
        'capCnt = i
    'End If
'End Sub

'=========================================================
' 预取所有“表格标题”段（带详细进度反馈）
' 输出：
'   capStarts()  每个命中段的 Range.Start
'   capTexts()   每个命中段的可见文本（已 Trim）
'   capKeys()    每个命中段所属的章键（由 headKeys/UpperBoundByStart 映射）
'   capCnt       命中总数
' 说明：
'   - 不改变原有输出语义，仅新增进度反馈，避免刷屏（抽样 + 节流）
'   - 进度条区间映射：62 → 70（与 BuildTableInfoArray 中的阶段区间错不开）
'=========================================================
Private Sub CollectAllCaptionParas(ByVal doc As Document, _
                                   ByRef headStarts() As Long, ByRef headKeys() As String, _
                                   ByRef capStarts() As Long, ByRef capTexts() As String, _
                                   ByRef capKeys() As String, ByRef capCnt As Long, _
                                   Optional ByRef pf As progressForm)

    Dim totalP As Long: totalP = doc.Paragraphs.Count
    Dim cap As Long, hit As Long
    Dim i As Long, tStep As Double, tLast As Double
    Dim px As Long, basePx As Long: basePx = 62
    Dim spanPx As Long: spanPx = 8         ' → 62..70

    '（一）初始化数组（按 1/64 文档段数预估，至少 64）
    cap = IIf(totalP \ 64 > 64, totalP \ 64, 64)
    ReDim capStarts(1 To cap)
    ReDim capTexts(1 To cap)
    ReDim capKeys(1 To cap)
    capCnt = 0

    '（二）缓存样式对象，失败则退化为名称比较
    Dim styCap As Style, safeByName As Boolean
    On Error Resume Next
    Set styCap = doc.Styles("表格标题")
    On Error GoTo 0
    safeByName = (styCap Is Nothing)

    '（三）开场提示
    If Not pf Is Nothing Then
        pf.UpdateProgressBar basePx, "②-预取‘表格标题’段：开始扫描（共 " & totalP & " 段）…"
        DoEvents
    End If

    tLast = Timer
    For i = 1 To totalP
        tStep = Timer

        Dim p As Paragraph
        Set p = doc.Paragraphs(i)

        ' 仅处理正文
        If p.Range.StoryType = wdMainTextStory Then
            ' 命中判定：优先对象比较，其次名称兜底
            Dim isCap As Boolean
            If Not safeByName Then
                On Error Resume Next
                isCap = (p.Range.Style Is styCap)
                If Err.Number <> 0 Then
                    Err.Clear
                    isCap = (CStr(p.Range.Style) = "表格标题")
                End If
                On Error GoTo 0
            Else
                isCap = (CStr(p.Range.Style) = "表格标题")
            End If

            If isCap Then
                ' 命中：扩容（倍增）
                hit = hit + 1
                If hit > cap Then
                    cap = cap * 2
                    ReDim Preserve capStarts(1 To cap)
                    ReDim Preserve capTexts(1 To cap)
                    ReDim Preserve capKeys(1 To cap)
                End If

                Dim st As Long: st = p.Range.Start
                capStarts(hit) = st
                capTexts(hit) = TrimVisible(FirstParaTextAtStart(doc, st))

                ' 所属章键：用传入的 headStarts/headKeys 二分上界
                Dim idx As Long: idx = UpperBoundByStart(headStarts, st)
                If idx >= 1 Then
                    capKeys(hit) = headKeys(idx)
                Else
                    capKeys(hit) = "0"
                End If
            End If
        End If

        ' ——（四）抽样 + 节流式进度反馈：每 1000 段 或 每 ~0.6s 回报一次
        If Not pf Is Nothing Then
            If (i Mod 1000 = 0) Or (Timer - tLast >= 0.6) Or (i = totalP) Then
                tLast = Timer
                px = basePx + CLng(spanPx * i / IIf(totalP = 0, 1, totalP))
                pf.UpdateProgressBar px, _
                    "②-预取进度：" & i & "/" & totalP & "  | 命中 " & hit & " 段"
                DoEvents
            End If
        End If
    Next i

    '（五）收尾：压缩到命中大小
    If hit = 0 Then
        Erase capStarts: Erase capTexts: Erase capKeys
        capCnt = 0
        If Not pf Is Nothing Then
            pf.UpdateProgressBar basePx + spanPx, "②-预取完成：未命中任何‘表格标题’段。"
        End If
        Exit Sub
    End If

    If hit < cap Then
        ReDim Preserve capStarts(1 To hit)
        ReDim Preserve capTexts(1 To hit)
        ReDim Preserve capKeys(1 To hit)
    End If
    capCnt = hit

    If Not pf Is Nothing Then
        pf.UpdateProgressBar basePx + spanPx, _
            "②-预取完成：共命中 " & capCnt & " 段。"
        DoEvents
    End If
End Sub


'============================== 工具函数（与前版一致） ==============================
Private Function SafeStyleName(ByVal r As Range) As String
    On Error Resume Next
    SafeStyleName = r.Style.NameLocal
    If Err.Number <> 0 Then SafeStyleName = ""
    On Error GoTo 0
End Function

Private Function SafeStyleNameByStart(ByVal pStart As Long) As String
    On Error Resume Next
    SafeStyleNameByStart = ActiveDocument.Range(pStart, pStart).Paragraphs(1).Range.Style.NameLocal
    If Err.Number <> 0 Then SafeStyleNameByStart = ""
    On Error GoTo 0
End Function

Private Function HeadingLevelByStyle(ByVal sty As String) As Long
    Select Case sty
        Case "标题 1", "标题1": HeadingLevelByStyle = 1
        Case "标题 2", "标题2": HeadingLevelByStyle = 2
        Case "标题 3", "标题3": HeadingLevelByStyle = 3
        Case "标题 4", "标题4": HeadingLevelByStyle = 4
        Case Else: HeadingLevelByStyle = 0
    End Select
End Function

Private Function SafeListString(ByVal r As Range) As String
    On Error Resume Next
    SafeListString = r.ListFormat.ListString
    If Err.Number <> 0 Then SafeListString = ""
    On Error GoTo 0
End Function

Private Function NormalizeChapterKey(ByVal listStr As String, ByVal lvl As Long) As String
    Dim s As String: s = Trim$(listStr)
    If s <> "" Then
        NormalizeChapterKey = s
    ElseIf lvl > 0 Then
        NormalizeChapterKey = "L" & CStr(lvl)
    Else
        NormalizeChapterKey = ""
    End If
End Function

Private Function PrevNonEmptyParaStart_ByStart(ByVal doc As Document, ByVal atStart As Long) As Long
    Dim prgs As Paragraphs
    On Error Resume Next
    Set prgs = doc.Range(0, atStart).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then
        PrevNonEmptyParaStart_ByStart = 0: Exit Function
    End If
    Dim p As Paragraph: Set p = prgs(prgs.Count)
    On Error GoTo 0
    Do While Not p Is Nothing
        Dim s As String: s = TrimVisible(p.Range.text)
        If s <> "" Then PrevNonEmptyParaStart_ByStart = p.Range.Start: Exit Function
        Set p = p.Previous
    Loop
    PrevNonEmptyParaStart_ByStart = 0
End Function

Private Function FirstParaTextAtStart(ByVal doc As Document, ByVal pStart As Long) As String
    If pStart <= 0 Then Exit Function
    Dim r As Range
    Set r = doc.Range(Start:=pStart, End:=doc.Range(pStart, pStart).Paragraphs(1).Range.End)
    FirstParaTextAtStart = TrimVisible(r.text)
End Function

Private Function TrimVisible(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    TrimVisible = Trim$(s)
End Function

Private Function UpperBoundByStart(ByRef starts() As Long, ByVal atStart As Long) As Long
    If (Not Not starts) = 0 Then Exit Function
    Dim lo As Long, hi As Long, mid As Long, ans As Long
    lo = 1
    hi = UBound(starts)
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If starts(mid) < atStart Then ans = mid: lo = mid + 1 Else hi = mid - 1
    Loop
    UpperBoundByStart = ans
End Function

'——（新增）HTML 转义：供本模块内部使用
Private Function HtmlEncode(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function


'（三补）生成“图片信息数组”（带进度：30%→100%）
'   定义沿用 1..12 列：
'   [1]=序号 [2]=图片起点 [3]=标题段起点（下方第一个非空段）
'   [4]=标题段文本 [5]=标题段样式 [6]=是否=【图片标题】
'   [7..11]=同表：最近标题级别/文本/编号/章键/章内序号
'   [12]=“图<章键>-<序号>”
Private Function BuildImageInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    Dim nInline As Long: nInline = doc.InlineShapes.Count
    Dim nShape As Long:  nShape = CountPictureShapes_InModule(doc)
    Dim total As Long:   total = nInline + nShape
    Dim baseW As Integer: baseW = 60

    Dim arr As Variant
    If total = 0 Then
        BuildImageInfoArray = Empty
        gOrphanRows = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "未检测到任何图片。"
        Exit Function
    End If

    ' ——收集所有图片的“文档位置起点”与“标题段起点”
    Dim pos() As Long, pStart() As Long, cnt As Long
    ReDim pos(1 To total): ReDim pStart(1 To total)

    Dim i As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        i = i + 1
        pos(i) = ils.Range.Start
        pStart(i) = NextNonEmptyParaStart_ByStart(doc, ils.Range.Start)
    Next

    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape_InModule(s) Then
            i = i + 1
            pos(i) = s.anchor.Start
            pStart(i) = NextNonEmptyParaStart_ByStart(doc, s.anchor.Start)
        End If
    Next
    cnt = i

    ' ——按位置排序（保证文档阅读顺序）
    Call SelectionSortByPos(pos, pStart, cnt)

    ' ——主数组
    ReDim arr(1 To cnt, 1 To 12)

    Dim dictSeq As Object:      Set dictSeq = CreateObject("Scripting.Dictionary")  ' 章键→章内序号
    Dim dictValidCap As Object: Set dictValidCap = CreateObject("Scripting.Dictionary") ' 有效标题段起点→True

    Dim k As Long
    For k = 1 To cnt
        If Not pf Is Nothing Then
            If (k Mod 10 = 0) Or (k = cnt) Then
                pf.UpdateProgressBar baseW + CInt((140# / IIf(cnt = 0, 1, cnt)) * k), "② 生成图片信息… " & k & "/" & cnt
            End If
        End If

        Dim atPos As Long: atPos = pos(k)
        Dim capStart As Long: capStart = pStart(k)

        ' 标题段文本/样式
        Dim pTxt As String, pSty As String
        If capStart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, capStart))
            pSty = SafeStyleNameByStart(capStart)
        Else
            pTxt = "": pSty = ""
        End If
        Dim isCap As Boolean: isCap = (pSty = "图片标题")
        If isCap And capStart > 0 Then dictValidCap(CStr(capStart)) = True

        ' 最近标题（上界）
        Dim idx As Long: idx = UpperBoundByStart(headStarts, atPos)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If

        ' 章内序号与“图号”
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "图" & key & "-" & CStr(seq)

        ' 写入（保持 1..12 列）
        arr(k, 1) = k
        arr(k, 2) = atPos
        arr(k, 3) = capStart
        arr(k, 4) = pTxt
        arr(k, 5) = pSty
        arr(k, 6) = isCap
        arr(k, 7) = lvl
        arr(k, 8) = hText
        arr(k, 9) = num
        arr(k, 10) = key
        arr(k, 11) = seq
        arr(k, 12) = label
    Next

    ' ——生成“孤儿段”（样式=图片标题，但未被任何图片命中）
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    CollectAllCaptionParas doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt, "图片标题"

    Dim orCnt As Long: orCnt = 0
    Dim orArr As Variant
    If capCnt > 0 Then
        Dim j As Long
        For j = 1 To capCnt
            Dim cs As Long: cs = capStarts(j)
            If cs <= 0 Then GoTo NEXT_J
            If Not dictValidCap.exists(CStr(cs)) Then
                Dim lvl2 As Long
                Dim idx2 As Long: idx2 = UpperBoundByStart(headStarts, cs)
                If idx2 >= 1 Then lvl2 = headLvls(idx2) Else lvl2 = 0
                orCnt = orCnt + 1
                If orCnt = 1 Then
                    ReDim orArr(1 To 3, 1 To 1)
                Else
                    ReDim Preserve orArr(1 To 3, 1 To orCnt)
                End If
                orArr(1, orCnt) = lvl2
                orArr(2, orCnt) = capTexts(j)
                orArr(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If

    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, t As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For t = 1 To orCnt
            outArr(t, 1) = orArr(1, t)
            outArr(t, 2) = orArr(2, t)
            outArr(t, 3) = orArr(3, t)
        Next
        gOrphanRows = outArr
    End If

    BuildImageInfoArray = arr
End Function

' ——辅助：当前位置后“下方第一个非空段”的起点
Private Function NextNonEmptyParaStart_ByStart(ByVal doc As Document, ByVal atStart As Long) As Long
    Dim prgs As Paragraphs
    On Error Resume Next
    Set prgs = doc.Range(atStart, doc.content.End).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then Exit Function
    Dim p As Paragraph: Set p = prgs(1).Next    ' 从“下一个段落”开始
    On Error GoTo 0
    Do While Not p Is Nothing
        Dim s As String: s = TrimVisible(p.Range.text)
        If s <> "" Then NextNonEmptyParaStart_ByStart = p.Range.Start: Exit Function
        Set p = p.Next
    Loop
    NextNonEmptyParaStart_ByStart = 0
End Function

' ——辅助：判断 Shape 是否为图片
Private Function IsPictureShape_InModule(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsPictureShape_InModule = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

' ——辅助：统计图片型浮动 Shape
Private Function CountPictureShapes_InModule(ByVal doc As Document) As Long
    Dim sh As Shape, N As Long
    For Each sh In doc.Shapes
        If IsPictureShape_InModule(sh) Then N = N + 1
    Next
    CountPictureShapes_InModule = N
End Function

' ——辅助：按文档位置升序排列（原地交换 pos、pstart 两个数组）
Private Sub SelectionSortByPos(ByRef pos() As Long, ByRef pStart() As Long, ByVal N As Long)
    Dim i As Long, j As Long, imin As Long, tp As Long, ts As Long
    For i = 1 To N - 1
        imin = i
        For j = i + 1 To N
            If pos(j) < pos(imin) Then imin = j
        Next
        If imin <> i Then
            tp = pos(i): pos(i) = pos(imin): pos(imin) = tp
            ts = pStart(i): pStart(i) = pStart(imin): pStart(imin) = ts
        End If
    Next
End Sub


'（一）计时器：开始
Private Function t0() As Double
    t0 = Timer
End Function

'（二）计时器：阶段耗时（并写入进度窗体）
Private Sub PF_Tick(ByVal pf As progressForm, ByRef t As Double, ByVal phase As String, Optional ByVal px As Long = -1)
    On Error Resume Next
    Dim dt As Double: dt = Timer - t: t = Timer
    If px < 0 Then
        pf.UpdateProgressBar pf.FrameProgress.width, "? " & phase & " 用时 " & Format(dt, "0.00") & " s"
    Else
        pf.UpdateProgressBar px, "? " & phase & " 用时 " & Format(dt, "0.00") & " s"
    End If
End Sub

'（三）心跳：不改变进度条位置，只追加一条消息
Private Sub PF_Log(ByVal pf As progressForm, ByVal msg As String)
    On Error Resume Next
    pf.UpdateProgressBar pf.FrameProgress.width, msg
End Sub


' 只在【耗时>阈值】且按抽样频率时写一条心跳（不强制每次输出，避免刷屏）
Private Sub PF_StepWarn(ByVal pf As progressForm, ByRef t As Double, _
                        ByVal tag As String, ByVal i As Long, ByVal N As Long, _
                        Optional ByVal warn As Double = 0.15, _
                        Optional ByVal sampleEvery As Long = 10, _
                        Optional ByVal basePx As Long = 70, _
                        Optional ByVal spanPx As Long = 80)
    On Error Resume Next
    Dim dt As Double: dt = Timer - t: t = Timer
    If (dt >= warn) And ((i <= 5) Or (i Mod sampleEvery = 0) Or (i = N)) Then
        Dim px As Long
        ' 将本阶段进度映射到 basePx~(basePx+spanPx)（与你入口条形进度区间一致）
        px = basePx + CLng(spanPx * i / IIf(N = 0, 1, N))
        pf.UpdateProgressBar px, "③ 表 " & i & "/" & N & " · " & tag & " 用时 " & Format(dt, "0.000") & " s"
        DoEvents
    End If
End Sub

