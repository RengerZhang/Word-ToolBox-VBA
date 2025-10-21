Attribute VB_Name = "图表标题_2_图表标题自动编号"
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

Public Sub 表格标题自动编号_使用进度窗体()
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


' ==========================================================
' 图片标题自动编号（带进度条；不删除手工编号；跳过表格内图片）
' 规则：
'   - 图号 = “图” + 最近标题号（优先四级→三级→二级→一级） + “-” + 同一第三级上下文内序号
'   - 显示号：有四级用四级 a.b.c.d，否则用现有级数（a / a.b / a.b.c）
'   - 序号分组键：固定按第三级 a.b.c（不足三级时按现有级数）
'   - 仅处理正文中的图片；且跳过位于表格内的图片
'   - 图片标题段：图片“下方的第一个非空段落”
'   - 若段首已存在“图 + 编号”前缀，则跳过（不覆盖、不删除）
'   - 进度窗体：ProgressForm（需已存在 UpdateProgressBar/stopFlag 等成员）
' ==========================================================
Public Sub 图片标题自动编号_使用进度窗体()
    Dim doc As Document: Set doc = ActiveDocument
    Dim totalPics As Long, done As Long
    Dim 图号计数 As Object: Set 图号计数 = CreateObject("Scripting.Dictionary")
'    Dim logPicShown As Long          ' 本次运行中已展示的示例条数
'    Const LOG_MAX As Long = 500        ' 最多展示 6 条“前后对照”

    
    ' ——统计“正文 & 非表格内”的图片数量（用于进度总数）
    totalPics = 统计正文非表格图片数量()
    
    ' ——打开进度窗体
    With progressForm
        .caption = "图片自动编号"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = "开始：共 " & totalPics & " 张图片（正文 & 非表格内）。" & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With
    
    If totalPics = 0 Then
        progressForm.UpdateProgressBar 200, "未发现符合条件的图片。"
        Unload progressForm
        MsgBox "未发现需要处理的图片（正文 & 非表格内）。", vbInformation
        Exit Sub
    End If
    
    ' ——收集图片位置并按文档顺序排序
    Dim pos() As Long, kind() As Integer, idx() As Long, cnt As Long
    收集图片位置_正文_非表格 doc, pos, kind, idx, cnt
    If cnt = 0 Then
        progressForm.UpdateProgressBar 200, "未发现符合条件的图片。"
        Unload progressForm
        Exit Sub
    End If
    排序图片位置 pos, kind, idx, cnt
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "图片自动编号"
    On Error GoTo 0
    
    ' ——遍历图片
    Dim i As Long, atStart As Long
    Dim capPara As Paragraph, capRange As Range
    Dim h As Range, 标题预览 As String
    Dim 段数组 As Variant, 显示号 As String, 计数Key As String
    Dim paraText As String, 正文 As String, 正文清 As String
    Const 图题样式名 As String = "图片标题"
    
    For i = 1 To cnt
        If progressForm.stopFlag Then Exit For
        
        atStart = pos(i)
            Dim logMsg As String
        
            '（1）定位“下方第一个非空段”作为图片标题段
            If kind(i) = 1 Then
                Set capPara = 下方首个非空段_从字符位置(doc, doc.InlineShapes(idx(i)).Range.End)
            Else
                Set capPara = 形状可视下方首段(doc, doc.Shapes(idx(i)))
            End If
            If capPara Is Nothing Then
                logMsg = "跳过：未找到图片下方的非空段。"
                GoTo REPORT
            End If
        
            '（2）若段首已有“图 + 编号”前缀 → 跳过
            paraText = 清理段首可见文本(capPara.Range.text)
            If 正则命中(paraText, "^\s*图\s*\d+(?:[\.．。]\s*\d+){0,6}\s*[-－–—]\s*\d+") Then
                logMsg = "跳过：已存在图号 → " & Left$(paraText, 80)
                GoTo REPORT
            End If
        
            '（3）可选：套用“图片标题”样式
            On Error Resume Next
            capPara.Style = doc.Styles("图片标题")
            On Error GoTo 0
        
            '（4）定位最近章节标题 → 解析显示号 & 分组 Key
            Set h = 定位最近标题_GoTo(capPara.Range)
            If Not h Is Nothing Then
                段数组 = 提取编号段数组(获取标准编号串(h.Paragraphs(1)))
                显示号 = 构造显示号_最多四级(段数组)
                计数Key = 构造计数Key_按第三级(段数组)
            Else
                显示号 = "": 计数Key = "0"
            End If
            If Len(计数Key) = 0 Then 计数Key = "0"
        
            '（5）组内自增序号
            If Not 图号计数.exists(计数Key) Then 图号计数.Add 计数Key, 0
            图号计数(计数Key) = 图号计数(计数Key) + 1
        
            '（6）写入：覆盖式写标准前缀 —— “图<显示号>-<序号>  ” + 原文（去掉任何残留前缀）
            Set capRange = capPara.Range.Duplicate
            If capRange.Characters.Count > 1 Then capRange.MoveEnd wdCharacter, -1  ' 去段尾标记
            正文 = capRange.text
            正文清 = 正则替换_一次(正文, "^\s*图\s*\d+(?:[\.．。]\s*\d+){0,6}\s*[-－–—]\s*\d+\s*", "")
            正文清 = LTrim$(正文清)
            capRange.text = "图" & 显示号 & "-" & CStr(图号计数(计数Key)) & "  " & 正文清
        
            ' ——本轮日志（写入场景才有“前/后对照”）
            logMsg = "写入 → 图" & 显示号 & "-" & 图号计数(计数Key) & vbCrLf & _
                     "  前：" & Left$(清理段首可见文本(正文), 80) & vbCrLf & _
                     "  后：" & Left$(清理段首可见文本(capRange.text), 80)
        
REPORT:
            done = done + 1
            progressForm.UpdateProgressBar 当前进度像素(done, totalPics), logMsg
            DoEvents
    Next i
    
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    If Not progressForm.stopFlag Then
        progressForm.UpdateProgressBar 200, "完成。"
        MsgBox "图片标题已写入图号（不删除手工编号；跳过表格内图片）。", vbInformation
    Else
        MsgBox "已手动中止。", vbExclamation
    End If
    
    On Error Resume Next
'    Unload progressForm
    On Error GoTo 0
End Sub


' =========================（辅助）图片枚举与定位 =========================

'（一）统计“正文 & 非表格内”的图片数量
Private Function 统计正文非表格图片数量() As Long
    Dim n As Long, ils As InlineShape, s As Shape
    For Each ils In ActiveDocument.InlineShapes
        If IsInlinePicture_Img(ils) Then
            If ils.Range.Paragraphs(1).Range.StoryType = wdMainTextStory Then
                If Not ils.Range.Information(wdWithInTable) Then n = n + 1
            End If
        End If
    Next
    For Each s In ActiveDocument.Shapes
        If IsFloatingPicture_Img(s) Then
            If s.anchor.Paragraphs(1).Range.StoryType = wdMainTextStory Then
                If Not s.anchor.Paragraphs(1).Range.Information(wdWithInTable) Then n = n + 1
            End If
        End If
    Next
    统计正文非表格图片数量 = n
End Function

'（二）收集图片文档起点（正文 & 非表格内），输出：pos/kind/idx，cnt 为元素数
'      kind: 1=InlineShape；2=Shape
Private Sub 收集图片位置_正文_非表格(ByVal doc As Document, _
    ByRef pos() As Long, ByRef kind() As Integer, ByRef idx() As Long, ByRef cnt As Long)

    Dim i As Long, ils As InlineShape, s As Shape
    ReDim pos(1 To doc.InlineShapes.Count + doc.Shapes.Count)
    ReDim kind(1 To UBound(pos))
    ReDim idx(1 To UBound(pos))
    
    ' Inline
    For i = 1 To doc.InlineShapes.Count
        Set ils = doc.InlineShapes(i)
        If IsInlinePicture_Img(ils) Then
            If ils.Range.Paragraphs(1).Range.StoryType = wdMainTextStory Then
                If Not ils.Range.Information(wdWithInTable) Then
                    cnt = cnt + 1
                    pos(cnt) = ils.Range.Start
                    kind(cnt) = 1
                    idx(cnt) = i
                End If
            End If
        End If
    Next i
    
    ' Shape
    For i = 1 To doc.Shapes.Count
        Set s = doc.Shapes(i)
        If IsFloatingPicture_Img(s) Then
            If s.anchor.Paragraphs(1).Range.StoryType = wdMainTextStory Then
                If Not s.anchor.Paragraphs(1).Range.Information(wdWithInTable) Then
                    cnt = cnt + 1
                    pos(cnt) = s.anchor.Start
                    kind(cnt) = 2
                    idx(cnt) = i
                End If
            End If
        End If
    Next i
    
    If cnt > 0 Then
        ReDim Preserve pos(1 To cnt)
        ReDim Preserve kind(1 To cnt)
        ReDim Preserve idx(1 To cnt)
    End If
End Sub

'（三）按起点升序排序（就地交换三组数组）
Private Sub 排序图片位置(ByRef pos() As Long, ByRef kind() As Integer, ByRef idx() As Long, ByVal n As Long)
    Dim i As Long, j As Long, imin As Long
    Dim tp As Long, tk As Integer, ti As Long
    For i = 1 To n - 1
        imin = i
        For j = i + 1 To n
            If pos(j) < pos(imin) Then imin = j
        Next
        If imin <> i Then
            tp = pos(i): pos(i) = pos(imin): pos(imin) = tp
            tk = kind(i): kind(i) = kind(imin): kind(imin) = tk
            ti = idx(i):  idx(i) = idx(imin):   idx(imin) = ti
        End If
    Next
End Sub

'（四）判断 Inline 是否为图片
Private Function IsInlinePicture_Img(ByVal ils As InlineShape) As Boolean
    On Error Resume Next
    Select Case ils.Type
        Case wdInlineShapePicture, wdInlineShapeLinkedPicture
            IsInlinePicture_Img = True
        Case Else
            IsInlinePicture_Img = False
    End Select
End Function

'（五）判断 Shape 是否为图片
Private Function IsFloatingPicture_Img(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsFloatingPicture_Img = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
End Function

'（六）找“下方第一个非空段”的 Paragraph（从起点所在段的下一段开始找）
Private Function 下方首个非空段_起点(ByVal doc As Document, ByVal atStart As Long) As Paragraph
    Dim prgs As Paragraphs, p As Paragraph
    On Error Resume Next
    Set prgs = doc.Range(atStart, doc.content.End).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then Exit Function
    Set p = prgs(1).Next                ' 下一段开始找
    On Error GoTo 0
    Do While Not p Is Nothing
        If Len(清理段首可见文本(p.Range.text)) > 0 Then
            Set 下方首个非空段_起点 = p
            Exit Function
        End If
        Set p = p.Next
    Loop
End Function



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
'（新）Inline：从“字符位置”开始，向下寻找第一个非空段（不含当前位置所在段的前半截）
Private Function 下方首个非空段_从字符位置(ByVal doc As Document, ByVal charPos As Long) As Paragraph
    Dim prgs As Paragraphs, p As Paragraph
    On Error Resume Next
    Set prgs = doc.Range(charPos, doc.content.End).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then Exit Function
    ' 从包含 charPos 的这一段“之后”的段落开始
    Set p = prgs(1).Next
    On Error GoTo 0
    Do While Not p Is Nothing
        If Len(清理段首可见文本(p.Range.text)) > 0 Then
            Set 下方首个非空段_从字符位置 = p
            Exit Function
        End If
        Set p = p.Next
    Loop
End Function

'（新）浮动 Shape：按“可视位置”在页面上找真正位于形状下方的首个非空段
' 实现思路：
'   （一）暂存当前选择 → 选择形状，读取其所在页号与页面内垂直位置（pt）
'   （二）自锚点所在段向后枚举段落，读取每个段落的“页号/页面内Top”
'   （三）第一个满足【页号>形状页】或【页号=形状页 且 段Top ≥ 形状Bottom】且非空 的段，即为目标
Private Function 形状可视下方首段(ByVal doc As Document, ByVal s As Shape) As Paragraph
    On Error GoTo FAIL
    Dim savedSel As Range
    Set savedSel = Selection.Range.Duplicate

    ' ① 读形状在页面上的位置
    s.Select
    Dim shpPage As Long
    Dim shpTop As Single, shpBottom As Single
    shpPage = Selection.Information(wdActiveEndAdjustedPageNumber)
    shpTop = Selection.Information(wdVerticalPositionRelativeToPage)
    shpBottom = shpTop + s.Height   ' s.Height 为 pt

    ' ② 从锚点段开始向后找
    Dim p As Paragraph
    Set p = s.anchor.Paragraphs(1)
    Do While Not p Is Nothing
        Dim txt As String: txt = 清理段首可见文本(p.Range.text)

        ' 段落的页号/Top
        p.Range.Select
        Dim pPage As Long, pTop As Single
        pPage = Selection.Information(wdActiveEndAdjustedPageNumber)
        pTop = Selection.Information(wdVerticalPositionRelativeToPage)

        ' 满足“在形状下方”的条件并且非空
        If (pPage > shpPage Or (pPage = shpPage And pTop >= shpBottom)) Then
            If Len(txt) > 0 Then
                Set 形状可视下方首段 = p
                Exit Do
            End If
        End If

        Set p = p.Next
    Loop

CLEAN:
    On Error Resume Next
    savedSel.Select    ' 还原选择，避免可见闪烁
    On Error GoTo 0
    Exit Function
FAIL:
    Resume CLEAN
End Function

