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

    Application.ScreenUpdating = False

    ' 2) 建索引（≈30%）
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT

    ' 3) 生成表数组（≈70%）
    arrTableInfo = BuildTableInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT

    Application.ScreenUpdating = True

    ' 4) 交给窗体
    Dim frm As 自检报告
    Set frm = New 自检报告
    frm.LoadReportFromArray arrTableInfo
    frm.Show vbModeless

    ' 5) 完成
    pf.UpdateProgressBar BAR_MAX, "完成。"
    DoEvents
ABORT:
    Unload pf
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
Private Function BuildTableInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    Dim n As Long: n = doc.Tables.Count
    Dim arr As Variant

    '（一）无表格时先标记返回空数组，但仍会继续收集孤儿段
    If n = 0 Then
        BuildTableInfoArray = Empty
        gOrphanRows = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "未检测到任何表格。"
        ' 即使无表格，下面仍会收集孤儿段（见后续 Collect）
    End If

    '（二）收集“样式=表格标题”的所有段：位置/文本/所属章键
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    CollectAllCaptionParas doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt

    '（三）逐表构建主数组，并记录“有效的表前标题段”
    If n > 0 Then ReDim arr(1 To n, 1 To 12)
    Dim dictSeq As Object:        Set dictSeq = CreateObject("Scripting.Dictionary")      ' 章键→章内序号
    Dim dictValidCap As Object:   Set dictValidCap = CreateObject("Scripting.Dictionary") ' 表前标题的起点→True

    Dim i As Long, tbl As Table
    Dim baseW As Integer: baseW = 60

    For i = 1 To n
        If Not pf Is Nothing Then
            If (i Mod 10 = 0) Or (i = n) Then
                pf.UpdateProgressBar baseW + CInt((140# / IIf(n = 0, 1, n)) * i), "② 生成表信息… " & i & "/" & n
            End If
        End If

        Set tbl = doc.Tables(i)
        Dim tStart As Long: tStart = tbl.Range.Start

        ' 1) 就近上一非空段
        Dim pstart As Long: pstart = PrevNonEmptyParaStart_ByStart(doc, tStart)
        Dim pTxt As String, pSty As String
        If pstart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, pstart))
            pSty = SafeStyleNameByStart(pstart)
        Else
            pTxt = "": pSty = ""
        End If

        ' 记录“有效标题”（确实位于表前的表格标题）
        If (pstart > 0) And (pSty = "表格标题") Then
            dictValidCap(CStr(pstart)) = True
        End If

        ' 2) 最近标题（上界二分）
        Dim idx As Long: idx = UpperBoundByStart(headStarts, tStart)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If

        ' 3) 章内序号与表号
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "表" & key & "-" & CStr(seq)

        ' 4) 是否命中【表格标题】
        Dim isCap As Boolean: isCap = (pSty = "表格标题")

        ' 5) 写入数组（保持原 1~12 列定义不变）
        arr(i, 1) = i
        arr(i, 2) = tStart
        arr(i, 3) = pstart
        arr(i, 4) = pTxt
        arr(i, 5) = pSty
        arr(i, 6) = isCap
        arr(i, 7) = lvl
        arr(i, 8) = hText
        arr(i, 9) = num
        arr(i, 10) = key
        arr(i, 11) = seq
        arr(i, 12) = label
    Next

    '（四）独立生成“孤儿段”二维数组：每行 [lvl, text, startPos]
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
                If idx2 >= 1 Then
                    lvl2 = headLvls(idx2)
                Else
                    lvl2 = 0
                End If

                orCnt = orCnt + 1
                ' ——按“最后一维可变”的方式扩容：orArr(3, 行)
                If orCnt = 1 Then
                    ReDim orArr(1 To 3, 1 To 1)
                Else
                    ReDim Preserve orArr(1 To 3, 1 To orCnt)   ' 只能改最后一维
                End If
                ' 写入当前行（注意下标次序改变）
                orArr(1, orCnt) = lvl2
                orArr(2, orCnt) = capTexts(j)
                orArr(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If        ' ←←← 这里是缺失的 End If，用于闭合 If capCnt > 0 Then

    '（五）输出给窗体使用的 gOrphanRows（转回 行×3 形状）
    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, k As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For k = 1 To orCnt
            outArr(k, 1) = orArr(1, k)   ' lvl
            outArr(k, 2) = orArr(2, k)   ' text
            outArr(k, 3) = orArr(3, k)   ' start
        Next
        gOrphanRows = outArr
    End If

    BuildTableInfoArray = arr
End Function





'——（四）工具：收集文档中所有样式=【表格标题】的段落（位置、文本、所属章键）
Private Sub CollectAllCaptionParas(ByVal doc As Document, _
                                   ByRef headStarts() As Long, ByRef headKeys() As String, _
                                   ByRef capStarts() As Long, ByRef capTexts() As String, _
                                   ByRef capKeys() As String, ByRef capCnt As Long, _
                                   Optional ByVal styleName As String = "表格标题")

    Dim total As Long: total = doc.Paragraphs.Count
    ReDim capStarts(1 To total)
    ReDim capTexts(1 To total)
    ReDim capKeys(1 To total)
    Dim p As Paragraph, i As Long

    For Each p In doc.Paragraphs
        If SafeStyleName(p.Range) = styleName Then
            i = i + 1
            capStarts(i) = p.Range.Start
            capTexts(i) = TrimVisible(FirstParaTextAtStart(doc, p.Range.Start))
            Dim idx As Long: idx = UpperBoundByStart(headStarts, p.Range.Start)
            If idx >= 1 Then
                capKeys(i) = headKeys(idx)
            Else
                capKeys(i) = "0"
            End If
        End If
    Next
    If i = 0 Then
        ReDim capStarts(1 To 1): capStarts(1) = 0
        ReDim capTexts(1 To 1):  capTexts(1) = ""
        ReDim capKeys(1 To 1):   capKeys(1) = "0"
        capCnt = 0
    Else
        ReDim Preserve capStarts(1 To i)
        ReDim Preserve capTexts(1 To i)
        ReDim Preserve capKeys(1 To i)
        capCnt = i
    End If
End Sub

'============================== 工具函数（与前版一致） ==============================
Private Function SafeStyleName(ByVal r As Range) As String
    On Error Resume Next
    SafeStyleName = r.Style.nameLocal
    If Err.Number <> 0 Then SafeStyleName = ""
    On Error GoTo 0
End Function

Private Function SafeStyleNameByStart(ByVal pstart As Long) As String
    On Error Resume Next
    SafeStyleNameByStart = ActiveDocument.Range(pstart, pstart).Paragraphs(1).Range.Style.nameLocal
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

Private Function FirstParaTextAtStart(ByVal doc As Document, ByVal pstart As Long) As String
    If pstart <= 0 Then Exit Function
    Dim r As Range
    Set r = doc.Range(Start:=pstart, End:=doc.Range(pstart, pstart).Paragraphs(1).Range.End)
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
    Dim pos() As Long, pstart() As Long, cnt As Long
    ReDim pos(1 To total): ReDim pstart(1 To total)

    Dim i As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        i = i + 1
        pos(i) = ils.Range.Start
        pstart(i) = NextNonEmptyParaStart_ByStart(doc, ils.Range.Start)
    Next

    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape_InModule(s) Then
            i = i + 1
            pos(i) = s.anchor.Start
            pstart(i) = NextNonEmptyParaStart_ByStart(doc, s.anchor.Start)
        End If
    Next
    cnt = i

    ' ——按位置排序（保证文档阅读顺序）
    Call SelectionSortByPos(pos, pstart, cnt)

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
        Dim capStart As Long: capStart = pstart(k)

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
    Dim sh As Shape, n As Long
    For Each sh In doc.Shapes
        If IsPictureShape_InModule(sh) Then n = n + 1
    Next
    CountPictureShapes_InModule = n
End Function

' ——辅助：按文档位置升序排列（原地交换 pos、pstart 两个数组）
Private Sub SelectionSortByPos(ByRef pos() As Long, ByRef pstart() As Long, ByVal n As Long)
    Dim i As Long, j As Long, imin As Long, tp As Long, ts As Long
    For i = 1 To n - 1
        imin = i
        For j = i + 1 To n
            If pos(j) < pos(imin) Then imin = j
        Next
        If imin <> i Then
            tp = pos(i): pos(i) = pos(imin): pos(imin) = tp
            ts = pstart(i): pstart(i) = pstart(imin): pstart(imin) = ts
        End If
    Next
End Sub


