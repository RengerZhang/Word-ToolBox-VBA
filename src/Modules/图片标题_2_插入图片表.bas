Attribute VB_Name = "图片标题_2_插入图片表"
Option Explicit

'========================
' 插入单栏图片表（图片内容控件版）
' 1) 光标处插入 1×2 表格，套“图片定位表”
' 2) 表格宽度自适应
' 3) 第1行插入【图片内容控件】，点击即可选图；行高按 4:3 预留可视空间
' 4) 第2行写入“图XXX  请在此处录入图名”，段落套“图片标题”
'========================
Public Sub 插入单栏图片表_图片控件版()
    '（一）样式名
    Const STYLE_PIC_TABLE As String = "图片定位表"
    Const PARA_STYLE_CAPTION As String = "图片标题"
    Const CC_TAG As String = "PIC_4TO3"

    '（二）准备与样式兜底
    Dim doc As Document: Set doc = ActiveDocument
    EnsureTableStyleOnly doc, STYLE_PIC_TABLE

    '（三）插入 1×2 表格 → 样式 → 自适应
    Dim rng As Range: Set rng = Selection.Range
    Dim tb As Table: Set tb = doc.Tables.Add(Range:=rng, NumRows:=2, NumColumns:=1)
    On Error Resume Next
    tb.Style = STYLE_PIC_TABLE
    On Error GoTo 0
    tb.AutoFitBehavior wdAutoFitWindow


    '（五）第1行：插入“图片内容控件”（点击弹出选图对话框）
    Dim cc As ContentControl
    tb.cell(1, 1).Range.text = ""       ' 清空单元格内容
    Set cc = doc.ContentControls.Add(wdContentControlPicture, tb.cell(1, 1).Range)

    With cc
        .Title = "图片（点击插入）"
        .tag = CC_TAG                   ' 便于后续批量处理
        .Appearance = wdContentControlBoundingBox
        ' 注：图片控件的占位提示是图标；也可附加一行提示文本：
        .SetPlaceholderText , , "点击此处插入图片（如果插入剪贴板内容，请选中此按钮并CTRL+V粘贴）"
    End With
    ' 居中显示控件
    cc.Range.ParagraphFormat.alignment = wdAlignParagraphCenter

    
    ' 第1行高度自动：随插入的图片大小自适应
    With tb.rows(1)
        .HeightRule = wdRowHeightAuto     ' 行高自动
    End With


    '（六）第2行：标题与样式
    With tb.cell(2, 1).Range
        .text = 构造图片标题占位_按新流程(ActiveDocument, tb.Range) & " 请在此处录入图名"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = PARA_STYLE_CAPTION
        On Error GoTo 0
    End With


    '（七）整体垂直居中（更美观）
    On Error Resume Next
    tb.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    On Error GoTo 0

    '（八）定位到表格后
    tb.Range.Select
    Selection.Collapse wdCollapseEnd
End Sub

'==============================
' 入口：插入“双栏图片定位表”（图上方两图，下一行两个子图名，最下总图名）
'==============================
Public Sub 插入双栏图片表_图片控件版_双栏()
    Dim doc As Document: Set doc = ActiveDocument
    Dim tb As Table, r As Long, c As Long
    Dim cc As ContentControl
    Dim 总图占位 As String
    
    '（一）先生成“总图名占位”（可用你的新流程函数）
    总图占位 = 构造图片标题占位_按新流程(doc, Selection.Range) & " 请在此处录入图名"
    
    '（二）在插入点建 3×2 表格
    Set tb = doc.Tables.Add(Selection.Range, 3, 2)
    
    '（三）套用表格样式 & 基础参数
    On Error Resume Next
    tb.Style = doc.Styles("图片定位表")
    On Error GoTo 0
    
    With tb
        .AllowAutoFit = True
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100                       ' 宽度自适应版心
        .rows.alignment = wdAlignRowCenter
        .rows.AllowBreakAcrossPages = False
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Spacing = 0
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
        .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
    
    '（四）第1行：两格用于放图 —— 行高自动；在单元格内插入“图片内容控件”
    With tb.rows(1)
        .HeightRule = wdRowHeightAuto               ' 关键：自动行高，随图自适应
    End With
    For c = 1 To 2
        With tb.cell(1, c).Range
            ' 单元格样式：图片格式
            On Error Resume Next
            .Style = ActiveDocument.Styles("图片格式")
            On Error GoTo 0
            ' 文字/对象居中
            .ParagraphFormat.alignment = wdAlignParagraphCenter
            ' 插入“图片”内容控件
            Set cc = .ContentControls.Add(wdContentControlPicture)
            cc.Title = IIf(c = 1, "图片a", "图片b")
            cc.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
            cc.SetPlaceholderText , , "单击此处插入图片"
        End With
    Next c
    
    '（五）第2行：左右子图名（a/b），样式=图片标题
    Call 导入图片标题_子图样式
    
    With tb.rows(2)
        .HeightRule = wdRowHeightAtLeast            ' 文本行给个最小高度更稳
        .Height = CentimetersToPoints(0.7)
    End With
    With tb.cell(2, 1).Range
        .text = "a） 输入子图名"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("图片标题-子图")
        On Error GoTo 0
    End With
    With tb.cell(2, 2).Range
        .text = "b） 输入子图名"
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("图片标题-子图")
        On Error GoTo 0
    End With
    
    '（六）第3行：合并为总图名，样式=图片标题
    tb.cell(3, 1).Merge tb.cell(3, 2)
    With tb.cell(3, 1).Range
        .text = 总图占位
        .ParagraphFormat.alignment = wdAlignParagraphCenter
        On Error Resume Next
        .Style = ActiveDocument.Styles("图片标题")
        On Error GoTo 0
    End With
    
    '（七）把光标收在表后，便于继续编辑
    tb.Range.Collapse wdCollapseEnd
    tb.Range.Select
End Sub


'========================
' 兜底：确保“表格样式”存在（不改外观）
'========================
Private Sub EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If
End Sub

'==========================================================
' 构造图片标题占位（按你最新版流程图）
' A = 最近标题(优先H4→H3→H2→H1) 的编号；若取不到则默认 "图1.1-1"
' B = 在上界=最近H3(退化到H2/H1/文首)，下界=从secStart向后最近的H1/H2/H3(无则文末)
'     区间内统计“图片标题”数量得到序号 idx = n + 1
'==========================================================
Private Function 构造图片标题占位_按新流程(ByVal doc As Document, ByVal atRng As Range) As String
    Dim chapA As String
    Dim anchorB As Paragraph
    Dim chapB As String
    Dim secStart As Long, secEnd As Long
    Dim n As Long, idx As Long

    '（一）A：最近标题编号（H4→H3→H2→H1）
    chapA = 取最近标题编号(atRng, Array(4, 3, 2, 1))
    If Len(chapA) = 0 Then
        构造图片标题占位_按新流程 = "图1.1-1"          ' 大纲未初始化 → 兜底
        Exit Function
    End If

    '（二）B 的上界：最近H3 → 无则H2 → 无则H1 → 无则文首
    Set anchorB = 取最近标题段落(atRng, Array(3, 2, 1))
    If anchorB Is Nothing Then
        secStart = doc.content.Start
        chapB = ""                                  ' 文首无编号
    Else
        secStart = anchorB.Range.End                ' 上界用 End（排除标题行）
        chapB = 安全取列表编号(anchorB)
    End If

    '（三）B 的下界：从 secStart 向后找 H1/H2/H3 三者中最先出现者；无则文末
    secEnd = 计算下界最早出现点(doc, secStart)

    '（四）统计区间 [secStart..secEnd) 内的“图片标题”数量（严格/稳妥二选一）
    n = 统计区间图片标题数(doc, secStart, secEnd, chapB)

    '（五）组装占位
    idx = n + 1
    构造图片标题占位_按新流程 = "图" & chapA & "-" & CStr(idx)
End Function

'==========================================================
' 取最近“符合给定级别集合”的标题编号（例如 levels = Array(4,3,2,1)）
' 找到即返回段落的 ListString；找不到返回 ""
'==========================================================
Private Function 取最近标题编号(ByVal atRng As Range, ByVal levels As Variant) As String
    Dim p As Paragraph
    Set p = 取最近标题段落(atRng, levels)
    If p Is Nothing Then
        取最近标题编号 = ""
    Else
        取最近标题编号 = 安全取列表编号(p)
    End If
End Function

'==========================================================
' 取最近标题“段落对象”（向前回溯；levels 例：Array(3,2,1)）
'==========================================================
Private Function 取最近标题段落(ByVal atRng As Range, ByVal levels As Variant) As Paragraph
    Dim p As Paragraph, prev As Paragraph, i As Long
    Set p = atRng.Paragraphs(1)

    Do While Not p Is Nothing
        For i = LBound(levels) To UBound(levels)
            If 段落是否指定级别标题(p, CLng(levels(i))) Then
                Set 取最近标题段落 = p
                Exit Function
            End If
        Next i
        Set p = 上一个段落(p)
    Loop
    Set 取最近标题段落 = Nothing
End Function

'==========================================================
' 判断段落是否指定级别标题（兼容中文“标题 n”与英文“Heading n”）
'==========================================================
Private Function 段落是否指定级别标题(ByVal p As Paragraph, ByVal lvl As Long) As Boolean
    On Error Resume Next
    Dim nm As String
    If TypeName(p.Range.Style) = "Style" Then
        nm = p.Range.Style.nameLocal
    Else
        nm = CStr(p.Range.Style)
    End If
    On Error GoTo 0

    nm = LCase$(nm)
    段落是否指定级别标题 = (nm = LCase$("标题 " & lvl)) Or (nm = LCase$("heading " & lvl))
End Function

'==========================================================
' 上一个段落（安全取法）
'==========================================================
'（二）上一个段落（Collapse + Move/Expand，不用 Duplicate）
Private Function 上一个段落(ByVal p As Paragraph) As Paragraph
    Dim r As Range
    Set r = p.Range.Duplicate ' 若你仍担心 Duplicate，这行可改为：Set r = p.Range

    r.Collapse wdCollapseStart
    If r.Start = 0 Then Exit Function               ' 文首
    r.MoveStart wdCharacter, -1                     ' 往前移 1 个字符
    r.Expand wdParagraph                            ' 扩展为上一个段落
    Set 上一个段落 = r.Paragraphs(1)
End Function



'==========================================================
' 安全取列表编号（ListString 可能取不到 → 返回 ""）
'==========================================================
Private Function 安全取列表编号(ByVal p As Paragraph) As String
    On Error Resume Next
    安全取列表编号 = Trim$(p.Range.ListFormat.ListString)
    If Err.Number <> 0 Then 安全取列表编号 = "": Err.Clear
    On Error GoTo 0
End Function

'==========================================================
' 计算下界：从 secStart 向后查 H1/H2/H3，取最先出现者的 Start；都无→文末
'==========================================================
Private Function 计算下界最早出现点(ByVal doc As Document, ByVal secStart As Long) As Long
    Dim p1 As Long, p2 As Long, p3 As Long
    p1 = 查找下一个标题起点(doc, secStart, 1)
    p2 = 查找下一个标题起点(doc, secStart, 2)
    p3 = 查找下一个标题起点(doc, secStart, 3)

    Dim m As Long: m = 极小正数(p1, p2, p3)
    If m = -1 Then
        计算下界最早出现点 = doc.content.End
    Else
        计算下界最早出现点 = m
    End If
End Function

'==========================================================
' 查找从 pos 向后“下一个 标题lvl”的起点；找不到返回 -1
'==========================================================
Private Function 查找下一个标题起点(ByVal doc As Document, ByVal pos As Long, ByVal lvl As Long) As Long
    Dim rng As Range: Set rng = doc.Range(Start:=pos, End:=doc.content.End)
    With rng.Find
        .ClearFormatting
        On Error Resume Next
        .Style = doc.Styles(IIf(lvl = 1, "标题 1", IIf(lvl = 2, "标题 2", "标题 3")))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    If rng.Find.Execute Then
        查找下一个标题起点 = rng.Start
        Exit Function
    End If
    ' 英文样式名兜底
    Set rng = doc.Range(Start:=pos, End:=doc.content.End)
    With rng.Find
        .ClearFormatting
        On Error Resume Next
        .Style = doc.Styles(IIf(lvl = 1, "Heading 1", IIf(lvl = 2, "Heading 2", "Heading 3")))
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    If rng.Find.Execute Then
        查找下一个标题起点 = rng.Start
    Else
        查找下一个标题起点 = -1
    End If
End Function

'==========================================================
' 返回三个位置的“最小正值”；若全为 -1（不可用）返回 -1
'==========================================================
Private Function 极小正数(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim t As Variant: t = Array(a, b, c)
    Dim i As Long, best As Long: best = -1
    For i = LBound(t) To UBound(t)
        If CLng(t(i)) >= 0 Then
            If best = -1 Or CLng(t(i)) < best Then best = CLng(t(i))
        End If
    Next i
    极小正数 = best
End Function

'==========================================================
' 统计区间内“图片标题”的数量（严格模式：限定前缀；若 chapB="" 则按样式计数）
'==========================================================
Private Function 统计区间图片标题数(ByVal doc As Document, ByVal startPos As Long, ByVal endPos As Long, ByVal chapB As String) As Long
    Dim scan As Range: Set scan = doc.Range(Start:=startPos, End:=endPos)
    Dim p As Paragraph, t As String, n As Long
    For Each p In scan.Paragraphs
        If 段落样式等于(p, "图片标题") Then
            t = 清理可见文本(p.Range.text)
            If Len(chapB) = 0 Then
                n = n + 1
            Else
                ' 匹配：^图<chapB>- 或 ^图<chapB>.
                If Left$(t, Len("图" & chapB & "-")) = "图" & chapB & "-" _
                   Or Left$(t, Len("图" & chapB & ".")) = "图" & chapB & "." Then
                    n = n + 1
                End If
            End If
        End If
    Next
    统计区间图片标题数 = n
End Function

Private Function 段落样式等于(ByVal p As Paragraph, ByVal styleName As String) As Boolean
    On Error Resume Next
    Dim nm As String
    If TypeName(p.Range.Style) = "Style" Then
        nm = p.Range.Style.nameLocal
    Else
        nm = CStr(p.Range.Style)
    End If
    On Error GoTo 0
    段落样式等于 = (LCase$(nm) = LCase$(styleName))
End Function

' 去掉回车/单元格结束符/空白
Private Function 清理可见文本(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    清理可见文本 = Trim$(s)
End Function

