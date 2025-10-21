Attribute VB_Name = "封面_一键设置封面"
Option Explicit
Public Const SIGN_MIN_LINESPACE_PT As Single = 4        ' 全工程可见

'===============================
' 封面生成（按你三点新要求）
'===============================
Public Sub 生成封面_绝对坐标()
    '——（一）文案（按需改）
    Dim 文_项目名 As String
    文_项目名 = "闵行区华漕镇 MHPO-1403 单元 73-04" & vbCrLf & "地块征收(动迁)安置住房项目"

    Dim 文_方案名 As String
    文_方案名 = "基坑支护、降水及土方开挖" & vbCrLf & "专项施工方案"

    ' 合并后的落款（单位+项目部）
    Dim 文_落款 As String
    文_落款 = "中交二公局东萌工程有限公司" & vbCrLf & _
             "闵行区华漕镇 MHPO-1403 单元 73-04 地块征收" & vbCrLf & "(动迁)安置住房项目经理部"

    Dim 文_日期 As String: 文_日期 = "2025年09月XX日"


    ' ——高度：均设为“> 两行”的安全值
    Const H_项目名_mm As Single = 30
    Const H_方案名_mm As Single = 35
    Const H_签字框_mm As Single = 50
    Const H_落款_mm  As Single = 35
    Const H_日期_mm   As Single = 12
    
     '——（二）版心/位置参数（沿用你原逻辑）
    Const 内边距_mm As Single = 0
    
    Dim y_mm As Single: y_mm = 25
    
    Dim Y_项目名_mm As Single: Y_项目名_mm = y_mm
    y_mm = y_mm + H_项目名_mm + 12
    
    Dim Y_方案名_mm As Single: Y_方案名_mm = y_mm
    y_mm = y_mm + H_方案名_mm + 25
    
    Dim Y_签字框_mm As Single: Y_签字框_mm = y_mm
    y_mm = y_mm + H_签字框_mm + 18
    
    Dim Y_落款_mm As Single: Y_落款_mm = y_mm
    y_mm = y_mm + H_落款_mm
    
    Dim Y_日期_mm As Single: Y_日期_mm = y_mm


    ' ——签字表左列宽
    Const W_签字框_左列_mm As Single = 30

    '（三）页面计算（版心宽高）
    Dim doc As Document: Set doc = ActiveDocument
    Dim ps As PageSetup: Set ps = doc.PageSetup
    Dim 可用宽 As Single, 可用高 As Single, 左 As Single, 上 As Single
    可用宽 = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    可用高 = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    左 = ps.LeftMargin + MM(内边距_mm)
    上 = ps.TopMargin + MM(内边距_mm)
    可用宽 = 可用宽 - 2 * MM(内边距_mm)
    可用高 = 可用高 - 2 * MM(内边距_mm)

    '（四）清理旧对象（删掉 DEPT）
    删除封面对象 "COVER_PROJ"
    删除封面对象 "COVER_TITLE"
    删除封面对象 "COVER_SIGNBOX"
    删除封面对象 "COVER_ORG"
    删除封面对象 "COVER_DATE"

    '（五）确保并更新两套独立样式（项目名、方案名）
    Ensure_Cover_Styles

    '（六）项目名（文本框：宽=版心，高>两行；样式=封面-项目名）
    放置文本框 tag:="COVER_PROJ", txt:=文_项目名, _
        x:=左 - 5, y:=上 + MM(Y_项目名_mm), w:=可用宽, h:=MM(H_项目名_mm), _
        applyStyle:="封面-项目名"

    '（七）方案名（文本框：宽=版心，高>两行；样式=封面-方案名）
    放置文本框 tag:="COVER_TITLE", txt:=文_方案名, _
        x:=左, y:=上 + MM(Y_方案名_mm), w:=可用宽 + MM(10), h:=MM(H_方案名_mm), _
        applyStyle:="封面-方案名"

    '（八）签字表（仍放在文本框里，第四点稍后讨论）
    放置签字表 "COVER_SIGNBOX", y:=上 + MM(Y_签字框_mm), w:=MM(70), h:=MM(40), _
        leftColWidthMM:=28, rightColWidthMM:=35, rowHeightMM:=12, _
        fontName:="宋体", fontPt:=14, bold:=True, _
        centerOnPage:=True


    '（九）落款（合并后的一段）
    放置文本框 tag:="COVER_ORG", txt:=文_落款, _
        x:=左, y:=上 + MM(Y_落款_mm), w:=可用宽, h:=MM(H_落款_mm), _
        applyStyle:="封面-落款", center:=True

    '（十）日期
    放置文本框 tag:="COVER_DATE", txt:=文_日期, _
        x:=左, y:=上 + MM(Y_日期_mm), w:=可用宽, h:=MM(H_日期_mm), _
        applyStyle:="封面-日期", center:=True

    MsgBox "封面生成完成（新版：样式化项目名/方案名，落款已合并）。", vbInformation
End Sub

'===============================
'（工具A）统一放置文本框（可选套“段落样式”）
'  - 若提供 applyStyle：优先使用该样式（含字体/行距），忽略 fontName/pt 设置
'  - 否则按传入的 fontName/fontPt/lineSpaceMultiple 设置
'===============================
Private Sub 放置文本框( _
    ByVal tag As String, ByVal txt As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
    Optional ByVal applyStyle As String = "", _
    Optional ByVal center As Boolean = True, _
    Optional ByVal fontName As String = "宋体", _
    Optional ByVal fontPt As Single = 12, _
    Optional ByVal bold As Boolean = False, _
    Optional ByVal lineSpaceMultiple As Single = 1.5)

    Dim shp As Shape
    Set shp = ActiveDocument.Shapes.AddTextBox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=x, Top:=y, width:=w, Height:=h, _
                anchor:=ActiveDocument.Range(0, 0))   ' 绝对定位

    With shp
        .line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .AlternativeText = tag
        .LockAnchor = True
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Left = wdShapeCenter
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Top = y

        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0:  .MarginBottom = 0
            .TextRange.text = txt

            If Len(applyStyle) > 0 Then
                On Error Resume Next
                .TextRange.Style = ActiveDocument.Styles(applyStyle)
                On Error GoTo 0
                ' 使用样式时，仅保证对齐；行距由样式控制
                .TextRange.ParagraphFormat.alignment = IIf(center, wdAlignParagraphCenter, wdAlignParagraphLeft)
            Else
                ' 直接设字体/行距
                With .TextRange.ParagraphFormat
                    .alignment = IIf(center, wdAlignParagraphCenter, wdAlignParagraphLeft)
                    .SpaceBefore = 0: .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = 12 * lineSpaceMultiple
                End With
                With .TextRange.Font
                    .NameFarEast = fontName
                    .NameAscii = fontName
                    .Size = fontPt
                    .bold = bold
                End With
            End If
        End With
    End With
End Sub

'===============================
'（工具B）放置签字表（保持你原实现）
'===============================
'===============================
' 放置签字表（文本框内嵌表格）
'  - centerOnPage=True：文本框相对“整页”水平+垂直居中
'  - 字体：宋体、加粗、四号（fontPt=14）
'  - 行距：1.5倍（段落行距）
'===============================
Private Sub 放置签字表( _
        ByVal tag As String, _
        Optional ByVal x As Single = 0, _
        Optional ByVal y As Single = 0, _
        Optional ByVal w As Single = 220, _
        Optional ByVal h As Single = 30, _
        Optional ByVal leftColWidthMM As Single = 30, _
        Optional ByVal rightColWidthMM As Single = 0, _
        Optional ByVal rowHeightMM As Single = 0, _
        Optional ByVal fontName As String = "宋体", _
        Optional ByVal fontPt As Single = 14, _
        Optional ByVal bold As Boolean = True, _
        Optional ByVal centerOnPage As Boolean = True _
        )
    
    Dim doc As Document: Set doc = ActiveDocument
    Dim ps As PageSetup: Set ps = Selection.Range.Sections(1).PageSetup
    
    
    ' ——创建承载表格的文本框（绝对定位到页面坐标）
    Dim shp As Shape
    Set shp = doc.Shapes.AddTextBox(msoTextOrientationHorizontal, _
        x, y, w, h, doc.Range(0, 0))
    With shp
        .AlternativeText = tag
        .line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .LockAnchor = True
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Top = y
    End With
    
    If centerOnPage Then
        shp.Left = wdShapeCenter        ' 相对于“页面”居中
    Else
        shp.Left = x                    ' 需要时仍可用绝对 x
    End If
    
    ' ——在文本框内部插入 3×2 表格
    shp.TextFrame.TextRange.text = ""
    shp.TextFrame.TextRange.Select
    Dim tb As Table
    Dim wPt As Single: wPt = w
    Dim leftPt As Single: leftPt = MM(leftColWidthMM)
    Dim rightPt As Single
    
    
    '左右边距的参数化定义
    If rightColWidthMM > 0 Then
        rightPt = MM(rightColWidthMM)
    Else
        rightPt = wPt - leftPt
    End If
    
    If leftPt + rightPt > wPt Then
        rightPt = wPt - leftPt
        If rightPt < 1 Then rightPt = 1
    End If
    
    
    Set tb = doc.Tables.Add(Selection.Range, 3, 2)
    
    With tb
        .AllowAutoFit = False
        .Borders.enable = False
        .rows.alignment = wdAlignRowCenter
        .rows.AllowBreakAcrossPages = False
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalBottom
        If rowHeightMM > 0 Then
            .rows.HeightRule = wdRowHeightExactly
            .rows.Height = MM(rowHeightMM)
        End If
        
        ' 列宽：左列固定，右列占满剩余宽度
        .Columns(1).width = leftPt
        .Columns(2).width = rightPt
        
        ' 左列文本
        .cell(1, 1).Range.text = "编  制："
        .cell(2, 1).Range.text = "审  核："
        .cell(3, 1).Range.text = "审  批："
        
        Dim r As Long
        For r = 1 To 3
            '（二-1）右列签名线：仅底边，**加粗**
            With .cell(r, 2).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt   ' ← 加粗
                .Color = wdColorAutomatic
            End With
        
            '（二-2）分别套用样式（左=右对齐；右=居中、段后0、1.5倍行距）
            .cell(r, 1).Range.Style = ActiveDocument.Styles("封面-签名左")
            .cell(r, 2).Range.Style = ActiveDocument.Styles("封面-签名右")
        Next r
    End With
End Sub

' mm → pt
Private Function MM(mmVal As Single) As Single
    MM = mmVal * 2.835
End Function


'===============================
'（工具C）创建/更新两套独立段落样式
'   - 封面-项目名：黑体/黑体，小一(24pt)，1.5倍行距，居中
'   - 封面-方案名：中文宋体，英文Times New Roman，小初(36pt)，1.5倍行距，居中
'===============================
Private Sub Ensure_Cover_Styles()
    Call EnsureParagraphStyle( _
        styleName:="封面-项目名", _
        nameCN:="黑体", nameEN:="黑体", _
        ptSize:=24, isBold:=False, lineRule:=wdLineSpaceSingle, align:=wdAlignParagraphCenter)

    Call EnsureParagraphStyle( _
        styleName:="封面-方案名", _
        nameCN:="宋体", nameEN:="Times New Roman", _
        ptSize:=36, isBold:=True, lineRule:=wdLineSpaceSingle, align:=wdAlignParagraphCenter)
        
        
    EnsureParagraphStyle "封面-落款", "宋体", "Times New Roman", 14, True, wdLineSpace1pt5, wdAlignParagraphCenter
    EnsureParagraphStyle "封面-日期", "宋体", "Times New Roman", 16, True, wdLineSpace1pt5, wdAlignParagraphCenter
    EnsureParagraphStyle "封面-签名左", "宋体", "Times New Roman", 14, True, SIGN_MIN_LINESPACE_PT, wdAlignParagraphRight
    EnsureParagraphStyle "封面-签名右", "仿宋", "Times New Roman", 15, False, SIGN_MIN_LINESPACE_PT, wdAlignParagraphCenter
    
    '（右列：去除底边距；两列统一 1.5 倍行距，段前后清零）
    With ActiveDocument.Styles("封面-签名左").ParagraphFormat
        .SpaceBefore = 0: .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = SIGN_MIN_LINESPACE_PT
    End With
    
    With ActiveDocument.Styles("封面-签名右").ParagraphFormat
        .SpaceBefore = 0: .SpaceAfter = 0   ' ← 右列去除底边距
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = SIGN_MIN_LINESPACE_PT
    End With
    

End Sub


'===============================
'（工具E）按 TAG 删除旧对象
'===============================
Private Sub 删除封面对象(tag As String)
    Dim s As Shape, i As Long
    On Error Resume Next
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        Set s = ActiveDocument.Shapes(i)
        If LCase$(s.AlternativeText) = LCase$(tag) Then s.Delete
    Next i
    On Error GoTo 0
End Sub


