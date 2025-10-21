Attribute VB_Name = "图表标题_1_图表标题样式统一"
Option Explicit

'==============================================
' 模块：图片与表标题匹配（带进度）
'==============================================

'（一）入口：图片标题样式统一（带进度）
'【要求】全文每张图片下方的第一个非空段落设置成“图片标题”样式
Public Sub 图片标题样式统一_带进度()
    Dim doc As Document: Set doc = ActiveDocument
    
    '（二）准备进度窗与目标样式
    Dim pf As progressForm
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "图片标题匹配（带进度）"
        pf.FrameProgress.width = 0
        pf.LabelPercentage.caption = "0%"
        pf.TextBoxStatus.text = "开始匹配图片标题……"
        pf.Show vbModeless
        DoEvents
    End If
    
    EnsureParaStyleExists doc, "图片标题"  ' 若不存在则创建（继承关系你在样式宏里已设，这里只保证有）
    
    '（三）计算总数（InlineShapes + 浮动Shapes中的图片）
    Dim total As Long, cntInline As Long, cntShape As Long
    cntInline = doc.InlineShapes.Count
    cntShape = CountPictureShapes(doc)
    total = cntInline + cntShape
    
    If total = 0 Then
        UpdateBar pf, 200, 200, "未发现任何图片。"
        GoTo CLEANUP
    End If
    
    '（四）逐个处理 InlineShapes（嵌入式图片）
    Dim i As Long, ils As InlineShape
    i = 0
    For Each ils In doc.InlineShapes
        i = i + 1
        If Not CaptionForInlineShape(doc, ils, "图片标题") Then
            StatusPulse pf, "（跳过）第 " & i & " 张 Inline 图片未找到下方非空段落。"
        End If
        
        UpdateBar pf, CInt(200# * i / total), 200, "处理图片（Inline）：" & i & "/" & total
        If Not pf Is Nothing Then If pf.stopFlag Then GoTo CLEANUP
    Next
    
    '（五）逐个处理浮动 Shapes（悬浮式图片）
    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape(s) Then
            i = i + 1
            If Not CaptionForShape(doc, s, "图片标题") Then
                StatusPulse pf, "（跳过）第 " & i & " 张浮动图片未找到下方非空段落。"
            End If
            
            UpdateBar pf, CInt(200# * i / total), 200, "处理图片（浮动）：" & i & "/" & total
            If Not pf Is Nothing Then If pf.stopFlag Then GoTo CLEANUP
        End If
    Next
    
    StatusPulse pf, "图片标题匹配完成：共处理 " & total & " 张图片。"

CLEANUP:
    If Not pf Is Nothing Then Unload pf
End Sub


'（六）入口：表标题样式统一（带进度）
' 逻辑与原“表标题匹配”一致：对每个表，向上寻找第一个非空段并设为【表格标题】样式
Public Sub 表标题样式统一_带进度()
    Dim doc As Document: Set doc = ActiveDocument
    
    Dim pf As progressForm
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "表标题匹配（带进度）"
        pf.FrameProgress.width = 0
        pf.LabelPercentage.caption = "0%"
        pf.TextBoxStatus.text = "开始匹配表标题……"
        pf.Show vbModeless
        DoEvents
    End If
    
    EnsureParaStyleExists doc, "表格标题"
    
    Dim n As Long: n = doc.Tables.Count
    If n = 0 Then
        UpdateBar pf, 200, 200, "文档中没有表格。"
        GoTo EXIT_B
    End If
    
    Dim i As Long, tbl As Table, rng As Range, prevPara As Paragraph
    For i = 1 To n
        Set tbl = doc.Tables(i)
        Set rng = tbl.Range: rng.Collapse wdCollapseStart
        Set prevPara = rng.Paragraphs(1).Previous
        
        Do While Not prevPara Is Nothing
            Dim t As String: t = 清理段落文本(prevPara.Range.text)
            If Len(t) > 0 Then
                prevPara.Style = doc.Styles("表格标题")
                Exit Do
            End If
            Set prevPara = prevPara.Previous
        Loop
        
        UpdateBar pf, CInt(200# * i / n), 200, "处理表格：" & i & "/" & n
        If Not pf Is Nothing Then If pf.stopFlag Then Exit For
    Next
    
    StatusPulse pf, "表标题匹配完成：共处理 " & n & " 张表。"

EXIT_B:
    If Not pf Is Nothing Then Unload pf
End Sub


'======================== 辅助函数区 ========================

'（七）对 InlineShape 设置“下方第一个非空段”为目标样式
Private Function CaptionForInlineShape(ByVal doc As Document, ByVal ils As InlineShape, ByVal styleName As String) As Boolean
    On Error GoTo SAFE_EXIT
    Dim p As Paragraph
    Set p = NextNonEmptyPara(ils.Range.Paragraphs(1))
    If Not p Is Nothing Then
        p.Style = doc.Styles(styleName)
        CaptionForInlineShape = True
    End If
SAFE_EXIT:
End Function

'（八）对浮动 Shape 设置“锚点下方第一个非空段”为目标样式
Private Function CaptionForShape(ByVal doc As Document, ByVal s As Shape, ByVal styleName As String) As Boolean
    On Error GoTo SAFE_EXIT
    If s Is Nothing Then Exit Function
    Dim anchorPara As Paragraph
    Set anchorPara = s.anchor.Paragraphs(1)
    Dim target As Paragraph
    Set target = NextNonEmptyPara(anchorPara)
    If Not target Is Nothing Then
        target.Style = doc.Styles(styleName)
        CaptionForShape = True
    End If
SAFE_EXIT:
End Function

'（九）从给定段落开始，向“下”寻找第一个非空段（不含当前段）
Private Function NextNonEmptyPara(ByVal p As Paragraph) As Paragraph
    Dim q As Paragraph
    If p Is Nothing Then Exit Function
    Set q = p.Next
    Do While Not q Is Nothing
        If Len(清理段落文本(q.Range.text)) > 0 Then
            Set NextNonEmptyPara = q
            Exit Function
        End If
        Set q = q.Next
    Loop
End Function

'（十）判断 Shape 是否为图片（msoPicture 或 msoLinkedPicture）
Private Function IsPictureShape(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsPictureShape = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

'（十一）统计文档中的“图片型”浮动 Shape 数量
Private Function CountPictureShapes(ByVal doc As Document) As Long
    Dim s As Shape, n As Long
    For Each s In doc.Shapes
        If IsPictureShape(s) Then n = n + 1
    Next
    CountPictureShapes = n
End Function

'（十二）保证段落样式存在（不存在则创建）
Private Sub EnsureParaStyleExists(ByVal doc As Document, ByVal styleName As String)
    On Error Resume Next
    Dim st As Style
    Set st = doc.Styles(styleName)
    If st Is Nothing Then
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
End Sub

'（十三）工具：清理段落可见文本（去尾标记/全角空格并 Trim）
Private Function 清理段落文本(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")                 ' 单元格结束符
    s = Replace$(s, ChrW(&H3000), " ")          ' 全角空格转半角
    清理段落文本 = Trim$(s)
End Function

'（十四）进度辅助：状态输出（不抢焦点）
Private Sub StatusPulse(ByVal pf As progressForm, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.TextBoxStatus.text = pf.TextBoxStatus.text & vbCrLf & msg
        pf.TextBoxStatus.SelStart = Len(pf.TextBoxStatus.text)
        pf.TextBoxStatus.SelLength = 0
        pf.Repaint
    End If
    DoEvents
    On Error GoTo 0
End Sub

'（十五）进度辅助：更新进度条（ProgressForm 的进度条总宽 200px）
Private Sub UpdateBar(ByVal pf As progressForm, ByVal cur As Long, ByVal total As Long, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.UpdateProgressBar cur, msg         ' 传入 0~200 的宽度
    End If
    DoEvents
    On Error GoTo 0
End Sub


