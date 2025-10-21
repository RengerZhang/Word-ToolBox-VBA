Attribute VB_Name = "编号_2_自动匹配各级标题"
Option Explicit

'==========================================================
' ② 正则表达式自动标题匹配（独立运行）
' 说明：
'   - 统一从 Mod配置中心 读取 pattern→style 的规则集
'   - 已编号段落优先：按 ListLevelNumber 映射到样式
'   - 未编号段落：用动态规则匹配段首编号形态，自动套用对应样式
'   - 不更改编号，只改变段落样式
'==========================================================
Sub 匹配标题并套用样式_基于配置中心()

    Dim doc As Document
    Dim rules As Variant ' [[pattern, style], ...]
    Dim cfg As Variant   ' 来自（③）“获取所有级别参数()”：[级, 列1..4]
    Dim level2Style() As String
    Dim Para As Paragraph
    Dim t As String
    Dim lvl As Long
    Dim i As Long
    Dim tocZones As Collection
    Set tocZones = 构建TOC区域集(doc)
    
    Set doc = ActiveDocument
    
    '——规则：从配置中心“按编号格式”动态生成（无需手写）
    rules = 生成标题匹配规则集()
    
    '——建立“编号级别 → 样式名”映射（供自动编号段落直接套样式）
    cfg = 获取所有级别参数()
    ReDim level2Style(1 To UBound(cfg, 1))
    For i = 1 To UBound(cfg, 1)
        level2Style(i) = CStr(cfg(i, 1)) ' 第1列：样式名
    Next i
    
     '——逐段处理
    For Each Para In doc.Paragraphs
    ' 0) 仅正文故事：排除页眉页脚、文本框等
    If Para.Range.StoryType <> wdMainTextStory Then GoTo NextPara

    ' 0.1) 排除“表格中的段落”（双保险：Information + Tables.Count）
    Dim inTable As Boolean
    On Error Resume Next
    inTable = Para.Range.Information(wdWithInTable)   ' True 表示在表格中
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If Not inTable Then inTable = (Para.Range.Tables.Count > 0)
    If inTable Then GoTo NextPara
    
    ' 0.2) 排除“目录中的段落”
    '     ① 样式名以 "TOC" 或 "目录" 开头
    '     ② 或者段落的 Range 落在任一 TOC 字段的结果区域内
    Dim sty As String
    On Error Resume Next
    sty = Para.Range.Style.nameLocal
    On Error GoTo 0
    If Len(sty) > 0 Then
        If (UCase$(Left$(sty, 3)) = "TOC" Or Left$(sty, 2) = "目录") Then GoTo NextPara
    End If
    If 在TOC区域内(Para.Range, tocZones) Then GoTo NextPara


    ' 1) 读取并清理可见文本
    t = 清理段落文本(Para.Range.text)
    If Len(t) = 0 Then GoTo NextPara

    ' 2) 自动编号优先：直接按级别映射
    On Error Resume Next
    If Para.Range.ListFormat.ListType <> wdListNoNumbering Then
        lvl = Para.Range.ListFormat.ListLevelNumber
    Else
        lvl = 0
    End If
    On Error GoTo 0

    If lvl >= LBound(level2Style) And lvl <= UBound(level2Style) Then
        If 样式存在(doc, level2Style(lvl)) Then
            Para.Style = doc.Styles(level2Style(lvl))
            GoTo NextPara
        End If
    End If

    ' 3) 非自动编号：用动态规则匹配段首编号模式 → 套样式
    If IsArray(rules) Then
        For i = LBound(rules, 1) To UBound(rules, 1)
            If 正则命中(t, CStr(rules(i, 1))) Then
                If 样式存在(doc, CStr(rules(i, 2))) Then
                    Para.Style = doc.Styles(CStr(rules(i, 2)))
                    Exit For
                End If
            End If
        Next i
    End If

NextPara:
Next Para
    
    MsgBox "②标题匹配完成！！", vbInformation
End Sub

'——工具：样式是否存在
Private Function 样式存在(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim s As Style
    On Error Resume Next
    Set s = doc.Styles(styleName)
    样式存在 = Not (s Is Nothing)
    Set s = Nothing
    On Error GoTo 0
End Function

'——工具：正则测试（仅判定）
Private Function 正则命中(ByVal s As String, ByVal pat As String) As Boolean
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    正则命中 = r.TEST(s)
End Function

'——工具：清理段落可见文本（去段尾标记/单元格结束符/全角空格→半角→Trim）
Private Function 清理段落文本(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")              ' 单元格结束符
    s = Replace$(s, ChrW(&H3000), " ")       ' 全角空格→半角
    清理段落文本 = Trim$(s)
End Function
'——构建 TOC 字段结果区域集合（结果 text 所在区间，而非域代码）
Private Function 构建TOC区域集(ByVal doc As Document) As Collection
    Dim zones As New Collection
    Dim f As Field, codeTxt As String
    On Error Resume Next
    For Each f In doc.Fields
        ' 用字段类型或代码文本判定（代码里关键字始终是 "TOC"）
        codeTxt = ""
        codeTxt = f.code.text
        If (f.Type = wdFieldTOC) Or (InStr(1, UCase$(codeTxt), "TOC", vbTextCompare) > 0) Then
            zones.Add f.Result.Duplicate
        End If
    Next f
    Set 构建TOC区域集 = zones
End Function

'——判定一个 Range 是否完全落在任意一个 TOC 结果区域中
Private Function 在TOC区域内(ByVal r As Range, ByVal zones As Collection) As Boolean
    Dim z As Range
    If zones Is Nothing Then Exit Function
    On Error Resume Next
    For Each z In zones
        If (r.Start >= z.Start) And (r.End <= z.End) Then
            在TOC区域内 = True
            Exit Function
        End If
    Next z
End Function


