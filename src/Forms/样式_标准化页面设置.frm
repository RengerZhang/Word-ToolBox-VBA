VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 样式_标准化页面设置 
   Caption         =   "中交标准化页面设置"
   ClientHeight    =   6150
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7210
   OleObjectBlob   =   "样式_标准化页面设置.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "样式_标准化页面设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'（一）数据结构定义
' === 用户设置结构体：存储表单输入的各项参数 ===
Private Type UserSettings
    topM As Double          ' 竖版顶边距
    bottomM As Double       ' 竖版底边距
    leftM As Double         ' 竖版左边距
    rightM As Double        ' 竖版右边距
    topM_L As Double        ' 横版顶边距
    bottomM_L As Double     ' 横版底边距
    leftM_L As Double       ' 横版左边距
    rightM_L As Double      ' 横版右边距
    headerLeft As String    ' 页眉左侧文本
    headerRight As String   ' 页眉右侧文本
    logoPath As String      ' Logo路径
    headerDist As Double    ' 页眉到页边距距离
    footerDist As Double    ' 页脚到页边距距离
End Type


'（二）辅助工具函数
' === 获取正文起始节索引：定位第一个一级大纲段落所在节 ===
' 若未找到一级大纲，则默认从第1节开始
Private Function 获取正文起始节索引() As Long
    Dim p As Paragraph
    获取正文起始节索引 = 1  ' 默认起始节为1
    For Each p In ActiveDocument.Paragraphs
        ' 通过大纲级别判断（比依赖样式名称更通用）
        If p.outlineLevel = wdOutlineLevel1 Then
            获取正文起始节索引 = p.Range.Sections(1).Index
            Exit Function
        End If
    Next p
End Function

' === 收集表单输入到UserSettings结构体 ===
Private Sub CollectSettings(ByRef settings As UserSettings)
    ' 从表单控件读取参数并转换为对应类型
    settings.topM = CDbl(txtTop.text)
    settings.bottomM = CDbl(txtBottom.text)
    settings.leftM = CDbl(txtLeft.text)
    settings.rightM = CDbl(txtRight.text)
    settings.topM_L = CDbl(txtTopL.text)
    settings.bottomM_L = CDbl(txtBottomL.text)
    settings.leftM_L = CDbl(txtLeftL.text)
    settings.rightM_L = CDbl(txtRightL.text)
    settings.headerLeft = txtHeaderLeft.text
    settings.headerRight = txtHeaderRight.text
    settings.logoPath = txtLogo.text
    settings.headerDist = CDbl(txtHeaderDist.text)
    settings.footerDist = CDbl(txtFooterDist.text)
End Sub

' === 统一字体设置：应用于页眉文本框和其他需要统一格式的内容 ===
Private Sub 统一字体(rng As Range)
    ' 1. 字体样式：宋体（中文）/Times New Roman（英文），10.5号，加粗，黑色
    With rng.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Color = wdColorBlack
        .bold = True
    End With
    
    ' 2. 段落格式：单倍行距，无段前段后间距，无首行缩进
    With rng.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
    End With
End Sub
' === 设置页眉样式：若存在则直接修改；不存在则创建 ===
Private Sub 设置页眉样式()
    Dim headerStyle As Style

    ' 1) 尝试获取已存在的样式
    On Error Resume Next
    Set headerStyle = ActiveDocument.Styles("HeaderStyle")
    On Error GoTo 0

    ' 2) 不存在则新建为“段落样式”
    If headerStyle Is Nothing Then
        Set headerStyle = ActiveDocument.Styles.Add( _
            name:="HeaderStyle", Type:=wdStyleTypeParagraph)
    Else
        ' 3) 已存在但不是“段落样式”，则重建为段落样式（避免类型冲突）
        If headerStyle.Type <> wdStyleTypeParagraph Then
            ' 如不希望提醒，直接删除重建即可
            headerStyle.Delete
            Set headerStyle = ActiveDocument.Styles.Add(name:="HeaderStyle", Type:=wdStyleTypeParagraph)
        End If
    End If

    ' 4) 统一样式格式
    On Error Resume Next
    headerStyle.AutomaticallyUpdate = False
    headerStyle.BaseStyle = ActiveDocument.Styles(wdStyleNormal) ' 以“正文/Normal”为基
    On Error GoTo 0

    With headerStyle.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Color = wdColorBlack
        .bold = True
    End With

    With headerStyle.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .alignment = wdAlignParagraphLeft    ' 左对齐
        ' 可选：避免同样式段落之间自动加空白（按需打开）
        ' .NoSpaceBetweenParagraphsOfSameStyle = True
    End With
End Sub


' === 应用页眉样式到页眉区域 ===
Private Sub 应用页眉样式到页眉(sec As Section)
    Dim hdr As HeaderFooter
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    ' 为页眉整体应用样式（文本框会继承基础格式）
    hdr.Range.Style = ActiveDocument.Styles("HeaderStyle")
End Sub


'（三）核心处理过程
' === 应用设置到单节：处理指定节的页边距、页眉、页脚、Logo等 ===
Private Sub ApplySettingsToSection(sec As Section, settings As UserSettings)
    Dim hdr As HeaderFooter, ftr As HeaderFooter  ' 页眉页脚对象
    Dim shp As Shape                               ' Logo图形对象
    Dim leftTextBox As Shape, rightTextBox As Shape ' 页眉左右文本框
    Dim frameWidth As Single                       ' 页眉容器宽度
    Dim 页眉底线顶边距 As Single                   ' 页眉底线的垂直位置
    Dim 文本框高度 As Single                       ' 页眉文本框高度
    
    ' 1. 页边距设置（区分横竖版）
    With sec.PageSetup
        If .Orientation = wdOrientPortrait Then  ' 竖版
            .TopMargin = CentimetersToPoints(settings.topM)
            .BottomMargin = CentimetersToPoints(settings.bottomM)
            .LeftMargin = CentimetersToPoints(settings.leftM)
            .RightMargin = CentimetersToPoints(settings.rightM)
        Else  ' 横版
            .TopMargin = CentimetersToPoints(settings.topM_L)
            .BottomMargin = CentimetersToPoints(settings.bottomM_L)
            .LeftMargin = CentimetersToPoints(settings.leftM_L)
            .RightMargin = CentimetersToPoints(settings.rightM_L)
        End If
        .HeaderDistance = CentimetersToPoints(settings.headerDist)
        .FooterDistance = CentimetersToPoints(settings.footerDist)
    End With
    
    ' 2. 页眉基础参数计算
    页眉底线顶边距 = CentimetersToPoints(1.5)  ' 页眉底线的垂直位置（固定1.5cm）
    文本框高度 = CentimetersToPoints(1)         ' 文本框高度固定1cm
    frameWidth = sec.PageSetup.PageWidth - sec.PageSetup.LeftMargin - sec.PageSetup.RightMargin  ' 可用宽度（页面宽-左右边距）
    
    ' 3. 页眉初始化（清空原有内容）
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    hdr.Range.text = ""
    
    ' 4. 应用页眉样式（确保文本框默认使用该样式）
    Call 设置页眉样式  ' 创建/更新样式
    Call 应用页眉样式到页眉(sec)  ' 应用到当前节页眉
    
    ' 5. 清除页眉默认边框（避免干扰）
    With hdr.Range.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    End With
    
    ' 6. 创建左侧文本框（关联页眉样式，输入时直接应用）
    Set leftTextBox = hdr.Shapes.AddTextBox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=0, Top:=0, width:=frameWidth * 0.45, Height:=文本框高度)  ' 占45%宽度
    With leftTextBox
        ' 定位与环绕设置
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .WrapFormat.Type = wdWrapNone  ' 不环绕文本
        .Left = sec.PageSetup.LeftMargin  ' 左对齐页边距
        .Top = 页眉底线顶边距 - 文本框高度 * 0.5  ' 垂直居中于底线上方
        
        ' 文本框格式（无边框，边距为0）
        .line.Visible = msoFalse
        .TextFrame.VerticalAnchor = msoAnchorBottom  ' 文本垂直靠下
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0
        .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        .TextFrame.AutoSize = False
    End With
    ' 填充左侧文本并应用样式（直接关联HeaderStyle，编辑时自动继承）
    With leftTextBox.TextFrame.TextRange
        .text = settings.headerLeft
        .Style = ActiveDocument.Styles("HeaderStyle")  ' 强制应用页眉样式
        .ParagraphFormat.alignment = wdAlignParagraphLeft  ' 左对齐
    End With
    
    ' 7. 创建右侧文本框（同上，关联样式）
    Set rightTextBox = hdr.Shapes.AddTextBox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=0, Top:=0, width:=frameWidth * 0.4, Height:=文本框高度)  ' 占40%宽度
    With rightTextBox
        ' 定位与环绕设置
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .WrapFormat.Type = wdWrapNone
        .Left = sec.PageSetup.LeftMargin + frameWidth * 0.6  ' 右移60%宽度（靠右）
        .Top = 页眉底线顶边距 - 文本框高度 * 0.5  ' 与左侧文本框垂直对齐
        
        ' 文本框格式（同左侧）
        .line.Visible = msoFalse
        .TextFrame.VerticalAnchor = msoAnchorBottom
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0
        .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        .TextFrame.AutoSize = False
    End With
    ' 填充右侧文本并应用样式
    With rightTextBox.TextFrame.TextRange
        .text = settings.headerRight
        .Style = ActiveDocument.Styles("HeaderStyle")  ' 强制应用页眉样式
        .ParagraphFormat.alignment = wdAlignParagraphRight  ' 右对齐
    End With
    
    ' 8. 页眉底线设置（蓝色单实线）
    With hdr.Range.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
            .Color = RGB(0, 72, 152)
        End With
        .Borders.DistanceFromTop = 1
        .Borders.DistanceFromLeft = 4
        .Borders.DistanceFromBottom = 1
        .Borders.DistanceFromRight = 4
        .Borders.Shadow = False
    End With
    
    ' 9. Logo处理（浮于文字上方，右下角定位到页边距）
    If settings.logoPath <> "" And Dir(settings.logoPath) <> "" Then
        Set shp = hdr.Shapes.AddPicture( _
            FileName:=settings.logoPath, _
            LinkToFile:=False, SaveWithDocument:=True)
        With shp
            .LockAspectRatio = msoTrue  ' 保持比例
            .Height = CentimetersToPoints(1)  ' 固定高度1cm
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .WrapFormat.Type = wdWrapFront  ' 浮于文字上方
            ' 右下角定位到（左边距，顶边距）
            .Left = sec.PageSetup.LeftMargin - .width
            .Top = 页眉底线顶边距 - 文本框高度 * 0.5  ' 与文本框垂直对齐
            .ZOrder msoBringToFront  ' 置于顶层
        End With
    End If
    
    ' 10. 处理页脚
    Call SetFooter(sec)
End Sub

' === 设置页脚：处理页码格式和编号规则 ===
Private Sub SetFooter(sec As Section)
    Dim ftr As HeaderFooter
    Dim i As Integer
    Dim firstTitleSection As Section
    Dim isHeading1Found As Boolean
    
    ' 1. 初始化页脚（清空内容）
    Set ftr = sec.Footers(wdHeaderFooterPrimary)
    ftr.Range.text = ""
    isHeading1Found = False  ' 标记是否找到一级标题
    
    ' 2. 查找第一个一级标题所在节（用于重新编号）
    For i = 1 To ActiveDocument.Sections.Count
        If ActiveDocument.Sections(i).Range.Paragraphs(1).Style = "Heading 1" Then
            Set firstTitleSection = ActiveDocument.Sections(i)
            isHeading1Found = True
            Exit For
        End If
    Next i
    
    ' 3. 设置页码编号规则
    If isHeading1Found Then
        ' 从第一个一级标题所在节开始重新编号（起始为1）
        With firstTitleSection.Footers(wdHeaderFooterPrimary).PageNumbers
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
    Else
        ' 无一级标题时从第1节开始编号
        With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
    End If
    
    ' 4. 设置页码格式（居中，Times New Roman，10.5号）
    With ftr.Range.Paragraphs(1).Range
        .Font.name = "Times New Roman"
        .Font.Size = 10.5
        .Font.Color = wdColorBlack
        .Fields.Add Range:=.Characters.Last, Type:=wdFieldPage  ' 插入页码域
        .ParagraphFormat.alignment = wdAlignParagraphCenter
    End With
End Sub


'（四）用户交互事件
' === 按钮：应用到本节 ===
Private Sub cmdApplySection_Click()
    Dim settings As UserSettings
    ' 收集表单设置并应用到当前选中节
    Call CollectSettings(settings)
    Call ApplySettingsToSection(Selection.Sections(1), settings)
    MsgBox "已应用到本节！", vbInformation
End Sub

' === 按钮：全文应用 ===
Private Sub cmdApplyAll_Click()
    Dim settings As UserSettings
    Dim secCount As Long, i As Long
    Dim 起始节 As Long
    Dim pf As progressForm  ' 进度窗体（假设已定义）
    Dim rsp As VbMsgBoxResult
    Dim rng As Range
    
    ' 1. 收集表单设置
    Call CollectSettings(settings)
    
    ' 2. 初始化参数（总节数、正文起始节）
    secCount = ActiveDocument.Sections.Count
    起始节 = 获取正文起始节索引()
    
    ' 3. 显示进度窗体
    样式_标准化页面设置.Hide
    Set pf = New progressForm
    pf.Show vbModeless
    pf.InitForPageSetting secCount, 起始节
    
    ' 4. 确认用户操作
    rsp = MsgBox("是否开始全文页眉页脚格式化？", vbQuestion + vbOKCancel, "确认开始")
    If rsp <> vbOK Then
        pf.UpdateProgressBar 0, "用户取消，未开始执行。"
        Exit Sub
    End If
    
    ' 5. 循环应用到所有节
    For i = 1 To secCount
        ' 检查是否需要停止
        If pf.stopFlag Then
            pf.UpdateProgressBar pf.FrameProgress.width, "操作已停止，正在退出..."
            Exit For
        End If
        
        ' 滚动到当前节首（提升可视化体验）
        Set rng = ActiveDocument.Sections(i).Range
        rng.Collapse wdCollapseStart
        rng.Select
        On Error Resume Next
        ActiveWindow.ScrollIntoView rng, True
        On Error GoTo 0
        
        ' 应用设置到当前节
        Call ApplySettingsToSection(ActiveDocument.Sections(i), settings)
        
        ' 更新进度
        pf.UpdateProgressBar CLng(200# * i / secCount), _
            "已经完成第 " & i & " 节页眉页脚设置，共 " & secCount & " 节"
        DoEvents
    Next i
    
    ' 6. 完成提示
    If Not pf.stopFlag Then
        pf.UpdateProgressBar 200, "已经完成全文页眉页脚设置！请点击【取消】退出。"
    End If

'    样式_标准化页面设置.Show
    
End Sub


' === 按钮：取消 ===
Private Sub cmdCancel_Click()
    Unload Me  ' 关闭表单
End Sub


Private Sub txtHeaderLeft_Change()

End Sub

' === 表单初始化：设置控件默认值 ===
Private Sub UserForm_Initialize()
    ' 竖版边距默认值
    txtTop.text = "2.5": txtBottom.text = "2.5": txtLeft.text = "3": txtRight.text = "3"
    ' 横版边距默认值
    txtTopL.text = "3": txtBottomL.text = "3": txtLeftL.text = "2.5": txtRightL.text = "2.5"
    ' 页眉文本默认值（左侧带换行）
    txtHeaderLeft.text = "闵行区华漕镇MHP0-1403单元" & Chr(13) & "73-04地块征收(动迁)安置住房项目"
    txtHeaderRight.text = "施工组织设计"
    ' 页眉页脚距离默认值
    txtHeaderDist.text = "1.5": txtFooterDist.text = "1.75"
    ' Logo默认路径（示例）
    txtLogo.text = "C:\Users\Tony Zhang\Desktop\logo.png"
End Sub

' === 浏览Logo：打开文件选择器选择图片 ===
Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择Logo图片"
        .Filters.Clear
        .Filters.Add "图片文件", "*.png; *.jpg; *.jpeg; *.bmp; *.gif"
        If .Show = -1 Then Me.txtLogo.text = .SelectedItems(1)  ' 选中后更新路径
    End With
End Sub
'==================== 旧窗体：中交标准化页面设置 ====================
' 说明：
' （一）这 4 个入口允许工具箱窗体“远程调用”旧窗体逻辑，改动极小；
' （二）host 传工具箱窗体（Me），若传入则优先把 host 的值同步到本窗体后再执行；
' （三）内部直接调用现有的按钮事件（*_Click），不改你旧代码。

' ——（0）把工具箱里的控件值同步到本窗体（私有小工具）——
Private Sub PS_CopyFromHost(ByVal host As Object)
    On Error Resume Next
    ' 竖版
    Me.txtTop.text = host.txtTop.text
    Me.txtBottom.text = host.txtBottom.text
    Me.txtLeft.text = host.txtLeft.text
    Me.txtRight.text = host.txtRight.text
    ' 横版
    Me.txtTopL.text = host.txtTopL.text
    Me.txtBottomL.text = host.txtBottomL.text
    Me.txtLeftL.text = host.txtLeftL.text
    Me.txtRightL.text = host.txtRightL.text
    ' 页眉/页脚/LOGO
    Me.txtHeaderLeft.text = host.txtHeaderLeft.text
    Me.txtHeaderRight.text = host.txtHeaderRight.text
    Me.txtLogo.text = host.txtLogo.text
    Me.txtHeaderDist.text = host.txtHeaderDist.text
    Me.txtFooterDist.text = host.txtFooterDist.text
    On Error GoTo 0
End Sub

'（一）公共入口：初始化（把 host 的值同步过来，不执行业务）
Public Sub PS_InitFromHost(ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
End Sub

'（二）公共入口：应用到本节（可选同步 host → 本窗体，随后复用你现有按钮逻辑）
Public Sub PS_ApplySection(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    ' 直接调用你现有的“应用到本节”按钮事件过程
    On Error Resume Next
    Call cmdApplySection_Click
    On Error GoTo 0
End Sub

'（三）公共入口：全文应用（可选同步 host → 本窗体，随后复用你现有按钮逻辑）
Public Sub PS_ApplyAll(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    On Error Resume Next
    Call cmdApplyAll_Click
    On Error GoTo 0
End Sub

'（四）公共入口：浏览 Logo（若传 host 先同步，再弹出你已有的选择逻辑）
Public Sub PS_BrowseLogo(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    On Error Resume Next
    Call cmdBrowse_Click
    On Error GoTo 0
End Sub
'==================== /旧窗体：中交标准化页面设置 ====================


