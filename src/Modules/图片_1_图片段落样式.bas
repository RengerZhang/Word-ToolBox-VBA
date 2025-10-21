Attribute VB_Name = "图片_1_图片段落样式"
'==============================================================
' 功能：统一全文图片的段落样式为【图片格式】
' 约束：样式在“样式导入模块”中定义；本模块只做调用
' 排除：页眉页脚等非正文 Story；不排除表格中的图片
' 反馈：使用 ProgressForm 显示进度，可终止
'==============================================================
Public Sub 统一图片段落样式_使用进度窗体()
    Dim doc As Document: Set doc = ActiveDocument
    Dim styPicPara As Style
    
    '（一）校验目标样式是否已导入（不创建，只调用）
    On Error Resume Next
    Set styPicPara = doc.Styles("图片格式")
    On Error GoTo 0
    If styPicPara Is Nothing Then
        MsgBox "未找到样式【图片格式】。" & vbCrLf & _
               "请先在【样式导入】中导入该样式后再运行。", vbExclamation, "图片段落样式统一"
        Exit Sub
    End If
    
    '（二）统计图片总数（只算正文 Story）
    Dim nInline As Long, nShape As Long, total As Long
    nInline = CountInlinePictures_MainStory(doc)
    nShape = CountFloatingPictures_MainStory(doc)
    total = nInline + nShape
    
    If total = 0 Then
        MsgBox "文档中未检测到任何图片（正文部分）。", vbInformation
        Exit Sub
    End If
    
    '（三）打开进度窗体
    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "统一图片段落样式（图片格式）"
    pf.FrameProgress.width = 0
    pf.LabelPercentage.caption = "0%"
    pf.TextBoxStatus.text = "共检测到图片：" & total & "（Inline=" & nInline & "，Floating=" & nShape & "）"
    pf.Show vbModeless
    DoEvents
    
    Application.ScreenUpdating = False
    
    Dim done As Long, changed As Long, unchanged As Long
    Dim p As Paragraph
    
    '（四）处理 InlineShapes（嵌入式）
    Dim ils As InlineShape
    For Each ils In doc.InlineShapes
        If pf.stopFlag Then GoTo EARLY_OUT
        If IsInlinePicture(ils) Then
            Set p = ils.Range.Paragraphs(1)
            ' 仅处理正文 Story
            If p.Range.StoryType = wdMainTextStory Then
                If Not ParaHasStyle(p, styPicPara) Then
                    p.Style = styPicPara
                    changed = changed + 1
                Else
                    unchanged = unchanged + 1
                End If
                done = done + 1
                pf.UpdateProgressBar ProgressPixels(done, total), "处理（Inline）：" & done & "/" & total
            End If
        End If
        DoEvents
    Next
    
    '（五）处理浮动 Shapes（悬浮图片）
    Dim s As Shape
    For Each s In doc.Shapes
        If pf.stopFlag Then GoTo EARLY_OUT
        If IsFloatingPicture(s) Then
            ' 取锚点所在段（可能在表格里——不排除）
            Set p = s.anchor.Paragraphs(1)
            If p.Range.StoryType = wdMainTextStory Then
                If Not ParaHasStyle(p, styPicPara) Then
                    p.Style = styPicPara
                    changed = changed + 1
                Else
                    unchanged = unchanged + 1
                End If
                done = done + 1
                pf.UpdateProgressBar ProgressPixels(done, total), "处理（浮动）：" & done & "/" & total
            End If
        End If
        DoEvents
    Next
    
EARLY_OUT:
    Application.ScreenUpdating = True
    
    If pf.stopFlag Then
        pf.UpdateProgressBar 200, "用户中止；已处理：" & done & "/" & total
        MsgBox "已中止。已处理：" & done & "/" & total & "；其中变更 " & changed & " 段。", vbExclamation
    Else
        pf.UpdateProgressBar 200, "完成。总计：" & total & "；变更 " & changed & "；已是目标样式 " & unchanged
        MsgBox "统一完成！" & vbCrLf & _
               "总图片数：" & total & vbCrLf & _
               "设置为【图片格式】：" & changed & vbCrLf & _
               "原本已为【图片格式】：" & unchanged, vbInformation
    End If
    
    Unload pf
End Sub

'=========================== 工具函数 ===========================

'（A）是否为图片（InlineShape）
Private Function IsInlinePicture(ByVal ils As InlineShape) As Boolean
    On Error Resume Next
    Select Case ils.Type
        Case wdInlineShapePicture, wdInlineShapeLinkedPicture
            IsInlinePicture = True
        Case Else
            IsInlinePicture = False
    End Select
End Function

'（B）是否为图片（Shape：浮动）
Private Function IsFloatingPicture(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsFloatingPicture = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
End Function

'（C）统计正文 Story 的 Inline 图片数量
Private Function CountInlinePictures_MainStory(ByVal doc As Document) As Long
    Dim n As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        If IsInlinePicture(ils) Then
            If ils.Range.Paragraphs(1).Range.StoryType = wdMainTextStory Then n = n + 1
        End If
    Next
    CountInlinePictures_MainStory = n
End Function

'（D）统计正文 Story 的浮动图片数量
Private Function CountFloatingPictures_MainStory(ByVal doc As Document) As Long
    Dim n As Long, s As Shape
    For Each s In doc.Shapes
        If IsFloatingPicture(s) Then
            If s.anchor.Paragraphs(1).Range.StoryType = wdMainTextStory Then n = n + 1
        End If
    Next
    CountFloatingPictures_MainStory = n
End Function

'（E）段落是否已是目标样式（用对象比较更稳）
Private Function ParaHasStyle(ByVal p As Paragraph, ByVal sty As Style) As Boolean
    On Error Resume Next
    ParaHasStyle = (p.Range.Style Is sty)
End Function

'（F）把“当前/总数”换算成 ProgressForm 的像素（0~200）
Private Function ProgressPixels(ByVal cur As Long, ByVal tot As Long) As Integer
    If tot <= 0 Then
        ProgressPixels = 0
    Else
        ProgressPixels = CInt(200# * cur / tot)
    End If
End Function


