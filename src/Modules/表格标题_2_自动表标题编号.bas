Attribute VB_Name = "表格标题_2_自动表标题编号"
Option Explicit

' ==========================================================
' 表格标题自动编号（带进度条；不删除手工编号）
' 规则：
'   - 表号 = “表” + 最近标题号（优先四级→三级→二级→一级） + “-” + 同一第三级上下文内序号
'   - 显示号：有四级用四级 a.b.c.d，否则用现有级数（a / a.b / a.b.c）
'   - 序号分组键：固定按第三级 a.b.c（不足三级时按现有级数）
'   - 仅处理正文中的表；表前可有多个空白段（向上找第一个非空段作为表题段）
'   - 若段首已存在“表 + 编号”前缀，则跳过（不覆盖、不删除）
'   - 进度窗体：ProgressForm（需已存在 UpdateProgressBar/stopFlag 等成员）
' ==========================================================

Public Sub 表格标题自动编号_使用进度窗体1()
    Dim doc As Document: Set doc = ActiveDocument
    Dim totalTables As Long, done As Long, passMsg As String
    Dim 表号计数 As Object: Set 表号计数 = CreateObject("Scripting.Dictionary")
    Dim tbl As Table, tblIdx As Long

    ' 统计正文中的表数量（用于进度条总数）
    totalTables = 统计正文表数量()

    ' 打开进度窗体
    With progressForm
        .caption = "表格自动编号"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = "开始：共 " & totalTables & " 个表。" & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "表格自动编号"
    On Error GoTo 0

    tblIdx = 0
    For Each tbl In doc.Tables
        If tbl.Range.StoryType <> wdMainTextStory Then GoTo NextTable
        tblIdx = tblIdx + 1
        If progressForm.stopFlag Then Exit For

        ' 编号处理（单个表）
        处理单个表 tbl, 表号计数, tblIdx, totalTables

        ' 进度
        done = tblIdx
        progressForm.UpdateProgressBar 当前进度像素(done, IIf(totalTables = 0, 1, totalTables)), _
            "进度：" & done & "/" & totalTables
        DoEvents

NextTable:
    Next tbl

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

    If Not progressForm.stopFlag Then
        progressForm.UpdateProgressBar 200, "完成。"
        MsgBox "表格标题已写入表号（不删除手工编号）。" & vbCrLf & _
               "提示：Ctrl+G 打开“立即窗口”查看详细日志。", vbInformation
    Else
        MsgBox "已手动中止。", vbExclamation
    End If

    On Error Resume Next
    Unload progressForm
    On Error GoTo 0
End Sub


' ----------------------------------------------------------
' 处理单个表：定位表题段→定位最近标题→生成显示号/计数Key→写入表号
' ----------------------------------------------------------
Private Sub 处理单个表(ByVal tbl As Table, _
                      ByRef 表号计数 As Object, _
                      ByVal 表序 As Long, _
                      ByVal totalTables As Long)

    Const 表题样式名 As String = "表格标题"

    Dim doc As Document: Set doc = ActiveDocument
    Dim tblRng As Range, prevPara As Paragraph, paraText As String
    Dim h As Range, 标题预览 As String
    Dim 级别 As Long, 原List As String, 解析号 As String
    Dim 段数组 As Variant, 显示号 As String, 计数Key As String
    Dim segDump As String
    Dim r As Range, 正文 As String, 正文清 As String

    Set tblRng = tbl.Range.Duplicate

    ' ——表前第一个非空段作为表题段（允许多个空白段）
    Set prevPara = 向上取第一个非空段(tblRng)
    If prevPara Is Nothing Then
        progressForm.UpdateProgressBar 当前进度像素(表序, IIf(totalTables = 0, 1, totalTables)), _
            "表#" & 表序 & "：未找到表题段，跳过。"
        Exit Sub
    End If

    ' ——套用“表格标题”样式（若已存在则不变）
    On Error Resume Next
    prevPara.Style = doc.Styles(表题样式名)
    On Error GoTo 0

    ' ——已有“表 + 编号”前缀则跳过（不覆盖）
    paraText = 清理段首可见文本(prevPara.Range.text)
    If 正则命中(paraText, "^\s*表\s*\d+(?:[\.．。]\s*\d+){0,6}\s*[-－–—]\s*\d+") Then
        progressForm.UpdateProgressBar 当前进度像素(表序, IIf(totalTables = 0, 1, totalTables)), _
            "表#" & 表序 & "：检测到已有表号，跳过。→ " & Left$(paraText, 40)
        Exit Sub
    End If

    ' ——定位最近章节标题（优先四级→三级→二级→一级）
    Set h = 定位最近标题_GoTo(prevPara.Range)
    If Not h Is Nothing Then
        级别 = h.Paragraphs(1).outlineLevel
        On Error Resume Next
        原List = h.Paragraphs(1).Range.ListFormat.ListString
        On Error GoTo 0

        解析号 = 获取标准编号串(h.Paragraphs(1))       ' 如 "3.1.4.1"
        段数组 = 提取编号段数组(解析号)                ' Array("3","1","4","1") 或 Empty
        显示号 = 构造显示号_最多四级(段数组)           ' 有 4 段用 4 段，否则用现有段
        计数Key = 构造计数Key_按第三级(段数组)         ' 始终按第三级分组
        标题预览 = 清理段首可见文本(h.Paragraphs(1).Range.text)
    Else
        级别 = 0: 原List = "": 解析号 = "": 显示号 = "": 计数Key = ""
        标题预览 = "(未找到标题)"
    End If

    ' ——调试输出
    If IsArray(段数组) Then
        On Error Resume Next
        segDump = Join(段数组, ",")
        If Err.Number <> 0 Then segDump = "(空)": Err.Clear
        On Error GoTo 0
    Else
        segDump = "(空)"
    End If

    Debug.Print "表#" & 表序 & "：级别=" & 级别 & _
                " | ListString=[" & 原List & "]" & _
                " | 解析号=[" & 解析号 & "]" & _
                " | 段数组=(" & segDump & ")" & _
                " | 显示号=[" & 显示号 & "]" & _
                " | 计数Key=[" & 计数Key & "]" & _
                " | 标题→ "; Left$(标题预览, 40)

    ' ——计数（同一第三级上下文内累加）
    If Len(计数Key) = 0 Then 计数Key = "0"
    If Not 表号计数.exists(计数Key) Then 表号计数.Add 计数Key, 0
    表号计数(计数Key) = 表号计数(计数Key) + 1

    ' ——写入表号（覆盖式：先清旧前缀，再写新前缀）
    Set r = prevPara.Range.Duplicate
    If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1

    正文 = r.text
    正文清 = 正则替换_一次(正文, "^\s*表\s*\d+(?:[\.．。]\s*\d+){0,6}\s*[-－–—]\s*\d+\s*", "")
    正文清 = LTrim$(正文清)

    r.text = "表" & 显示号 & "-" & CStr(表号计数(计数Key)) & "  " & 正文清

    progressForm.UpdateProgressBar 当前进度像素(表序, IIf(totalTables = 0, 1, totalTables)), _
        "表#" & 表序 & "：写入 → 表" & 显示号 & "-" & 表号计数(计数Key)
End Sub


' ========================= 辅助与工具 =========================

' 统计正文中表的数量
Private Function 统计正文表数量() As Long
    Dim t As Table
    Dim n As Long
    For Each t In ActiveDocument.Tables
        If t.Range.StoryType = wdMainTextStory Then n = n + 1
    Next
    统计正文表数量 = n
End Function

' 向上取第一个非空段（允许多个空白段）
Private Function 向上取第一个非空段(ByVal tblRng As Range) As Paragraph
    Dim p As Paragraph, s As String
    If tblRng.Paragraphs.Count = 0 Then Exit Function
    Set p = tblRng.Paragraphs(1).Previous
    Do While Not p Is Nothing
        s = 清理段首可见文本(p.Range.text)
        If Len(s) > 0 Then Set 向上取第一个非空段 = p: Exit Function
        Set p = p.Previous
    Loop
End Function

' 从锚点向上，按“就近+挡位”规则定位最近标题
' 优先四级→若遇到三级则优先最近四级，否则返回该三级；遇到二/一级直接返回
Private Function 定位最近标题_GoTo(ByVal anchor As Range) As Range
    Dim base As Range, cur As Range, hop As Range
    Dim cand4 As Range, lvl As Long, guard As Long

    Set base = anchor.Duplicate
    base.SetRange Start:=base.Start, End:=base.Start

    Set cur = base.Duplicate
    Do
        On Error Resume Next
        Set hop = cur.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        On Error GoTo 0
        If hop Is Nothing Then Exit Do
        If hop.Start >= cur.Start Then Exit Do  ' 防死循环

        Set cur = hop
        lvl = cur.Paragraphs(1).outlineLevel

        Select Case lvl
            Case wdOutlineLevel4
                If cand4 Is Nothing Then Set cand4 = cur.Paragraphs(1).Range
            Case wdOutlineLevel3
                If Not cand4 Is Nothing Then
                    Set 定位最近标题_GoTo = cand4
                Else
                    Set 定位最近标题_GoTo = cur.Paragraphs(1).Range
                End If
                Exit Function
            Case wdOutlineLevel2, wdOutlineLevel1
                Set 定位最近标题_GoTo = cur.Paragraphs(1).Range
                Exit Function
            Case Else
                ' 其他级别继续上跳
        End Select

        guard = guard + 1
        If guard > 20000 Then Exit Do
    Loop

    If Not cand4 Is Nothing Then Set 定位最近标题_GoTo = cand4
End Function

' 把标题段转为标准编号串（优先 ListString；失败则从段首文本解析）
Private Function 获取标准编号串(ByVal p As Paragraph) As String
    Dim s As String, t As String
    On Error Resume Next
    s = p.Range.ListFormat.ListString
    On Error GoTo 0
    s = 规范化编号串(s)
    If Len(s) > 0 Then
        获取标准编号串 = s
        Exit Function
    End If
    t = 解析段首编号(p.Range.text)
    获取标准编号串 = t
End Function

' 提取编号段数组：只保留数字与点，压缩多点，Split
Private Function 提取编号段数组(ByVal numStr As String) As Variant
    Dim s As String
    s = Replace$(Replace$(numStr, "．", "."), "。", ".")
    s = 正则替换_全局(s, "[^\d\.]", "")
    s = 正则替换_全局(s, "^\.+|\.+$", "")
    s = 正则替换_全局(s, "\.+", ".")
    If Len(s) = 0 Then
        提取编号段数组 = Empty
    Else
        提取编号段数组 = Split(s, ".")
    End If
End Function

' 显示号：有 4 段用 4 段；否则按现有段数（1/2/3）
Private Function 构造显示号_最多四级(ByVal segs As Variant) As String
    Dim n As Long
    If IsEmpty(segs) Then Exit Function
    n = UBound(segs) - LBound(segs) + 1
    Select Case n
        Case Is >= 4: 构造显示号_最多四级 = segs(0) & "." & segs(1) & "." & segs(2) & "." & segs(3)
        Case 3:       构造显示号_最多四级 = segs(0) & "." & segs(1) & "." & segs(2)
        Case 2:       构造显示号_最多四级 = segs(0) & "." & segs(1)
        Case Else:    构造显示号_最多四级 = segs(0)
    End Select
End Function

' 计数Key：固定用到第三级；不足三级时用现有段数
Private Function 构造计数Key_按第三级(ByVal segs As Variant) As String
    Dim n As Long
    If IsEmpty(segs) Then Exit Function
    n = UBound(segs) - LBound(segs) + 1
    Select Case n
        Case Is >= 3: 构造计数Key_按第三级 = segs(0) & "." & segs(1) & "." & segs(2)
        Case 2:       构造计数Key_按第三级 = segs(0) & "." & segs(1)
        Case Else:    构造计数Key_按第三级 = segs(0)
    End Select
End Function

' 规范化编号：去空白/全角点→半角；仅保留数字与点；压缩多点；去首尾点；基本校验
Private Function 规范化编号串(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace$(s, vbCr, "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = Replace$(s, "．", ".")
    s = Replace$(s, "。", ".")
    s = 正则替换_全局(s, "\s+", "")
    s = 正则替换_全局(s, "[^\d\.]", "")
    s = 正则替换_全局(s, "\.+", ".")
    s = 正则替换_全局(s, "^\.|\.?$", "")
    If 正则命中(s, "^\d+(?:\.\d+){0,7}$") Then 规范化编号串 = s
End Function

' 从段首文本解析编号（允许点前后空格/全角点）
Private Function 解析段首编号(ByVal s As String) As String
    Dim m As Object
    s = Replace$(Replace$(s, "．", "."), "。", ".")
    s = Replace$(s, vbCr, "")
    Set m = 正则匹配(s, "^\s*\d+(?:\s*\.\s*\d+){0,7}")
    If Not m Is Nothing Then 解析段首编号 = 规范化编号串(m.Value)
End Function

' 去段尾/单元格结束符/全角空格→半角，并 Trim
Private Function 清理段首可见文本(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    清理段首可见文本 = Trim$(s)
End Function

' ——正则工具
Private Function 正则命中(ByVal s As String, ByVal pat As String) As Boolean
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = False: rx.Global = False: rx.pattern = pat
    正则命中 = rx.TEST(s)
End Function

Private Function 正则匹配(ByVal s As String, ByVal pat As String) As Object
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    Dim mc As Object
    rx.IgnoreCase = False: rx.Global = False: rx.pattern = pat
    Set mc = rx.Execute(s)
    If mc.Count > 0 Then Set 正则匹配 = mc(0) Else Set 正则匹配 = Nothing
End Function

' 单次替换：用于删除已有“表+编号”前缀（只替首处，避免误删正文）
Private Function 正则替换_一次(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    正则替换_一次 = rx.Replace(s, rep)
End Function

' 全局替换
Private Function 正则替换_全局(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = False: rx.Global = True: rx.pattern = pat
    正则替换_全局 = rx.Replace(s, rep)
End Function

' 进度像素（你的窗体满条 200px）
Private Function 当前进度像素(ByVal done As Long, ByVal total As Long) As Long
    If total <= 0 Then
        当前进度像素 = 0
    Else
        当前进度像素 = CLng(200# * done / total)
        If 当前进度像素 < 0 Then 当前进度像素 = 0
        If 当前进度像素 > 200 Then 当前进度像素 = 200
    End If
End Function


