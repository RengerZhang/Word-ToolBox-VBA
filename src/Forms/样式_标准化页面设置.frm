VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��ʽ_��׼��ҳ������ 
   Caption         =   "�н���׼��ҳ������"
   ClientHeight    =   6150
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7210
   OleObjectBlob   =   "��ʽ_��׼��ҳ������.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��ʽ_��׼��ҳ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'��һ�����ݽṹ����
' === �û����ýṹ�壺�洢������ĸ������ ===
Private Type UserSettings
    topM As Double          ' ���涥�߾�
    bottomM As Double       ' ����ױ߾�
    leftM As Double         ' ������߾�
    rightM As Double        ' �����ұ߾�
    topM_L As Double        ' ��涥�߾�
    bottomM_L As Double     ' ���ױ߾�
    leftM_L As Double       ' �����߾�
    rightM_L As Double      ' ����ұ߾�
    headerLeft As String    ' ҳü����ı�
    headerRight As String   ' ҳü�Ҳ��ı�
    logoPath As String      ' Logo·��
    headerDist As Double    ' ҳü��ҳ�߾����
    footerDist As Double    ' ҳ�ŵ�ҳ�߾����
End Type


'�������������ߺ���
' === ��ȡ������ʼ����������λ��һ��һ����ٶ������ڽ� ===
' ��δ�ҵ�һ����٣���Ĭ�ϴӵ�1�ڿ�ʼ
Private Function ��ȡ������ʼ������() As Long
    Dim p As Paragraph
    ��ȡ������ʼ������ = 1  ' Ĭ����ʼ��Ϊ1
    For Each p In ActiveDocument.Paragraphs
        ' ͨ����ټ����жϣ���������ʽ���Ƹ�ͨ�ã�
        If p.outlineLevel = wdOutlineLevel1 Then
            ��ȡ������ʼ������ = p.Range.Sections(1).Index
            Exit Function
        End If
    Next p
End Function

' === �ռ������뵽UserSettings�ṹ�� ===
Private Sub CollectSettings(ByRef settings As UserSettings)
    ' �ӱ��ؼ���ȡ������ת��Ϊ��Ӧ����
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

' === ͳһ�������ã�Ӧ����ҳü�ı����������Ҫͳһ��ʽ������ ===
Private Sub ͳһ����(rng As Range)
    ' 1. ������ʽ�����壨���ģ�/Times New Roman��Ӣ�ģ���10.5�ţ��Ӵ֣���ɫ
    With rng.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Color = wdColorBlack
        .bold = True
    End With
    
    ' 2. �����ʽ�������о࣬�޶�ǰ�κ��࣬����������
    With rng.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = 0
        .SpaceAfter = 0
        .FirstLineIndent = 0
    End With
End Sub
' === ����ҳü��ʽ����������ֱ���޸ģ��������򴴽� ===
Private Sub ����ҳü��ʽ()
    Dim headerStyle As Style

    ' 1) ���Ի�ȡ�Ѵ��ڵ���ʽ
    On Error Resume Next
    Set headerStyle = ActiveDocument.Styles("HeaderStyle")
    On Error GoTo 0

    ' 2) ���������½�Ϊ��������ʽ��
    If headerStyle Is Nothing Then
        Set headerStyle = ActiveDocument.Styles.Add( _
            name:="HeaderStyle", Type:=wdStyleTypeParagraph)
    Else
        ' 3) �Ѵ��ڵ����ǡ�������ʽ�������ؽ�Ϊ������ʽ���������ͳ�ͻ��
        If headerStyle.Type <> wdStyleTypeParagraph Then
            ' �粻ϣ�����ѣ�ֱ��ɾ���ؽ�����
            headerStyle.Delete
            Set headerStyle = ActiveDocument.Styles.Add(name:="HeaderStyle", Type:=wdStyleTypeParagraph)
        End If
    End If

    ' 4) ͳһ��ʽ��ʽ
    On Error Resume Next
    headerStyle.AutomaticallyUpdate = False
    headerStyle.BaseStyle = ActiveDocument.Styles(wdStyleNormal) ' �ԡ�����/Normal��Ϊ��
    On Error GoTo 0

    With headerStyle.Font
        .NameFarEast = "����"
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
        .alignment = wdAlignParagraphLeft    ' �����
        ' ��ѡ������ͬ��ʽ����֮���Զ��ӿհף�����򿪣�
        ' .NoSpaceBetweenParagraphsOfSameStyle = True
    End With
End Sub


' === Ӧ��ҳü��ʽ��ҳü���� ===
Private Sub Ӧ��ҳü��ʽ��ҳü(sec As Section)
    Dim hdr As HeaderFooter
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    ' Ϊҳü����Ӧ����ʽ���ı����̳л�����ʽ��
    hdr.Range.Style = ActiveDocument.Styles("HeaderStyle")
End Sub


'���������Ĵ������
' === Ӧ�����õ����ڣ�����ָ���ڵ�ҳ�߾ࡢҳü��ҳ�š�Logo�� ===
Private Sub ApplySettingsToSection(sec As Section, settings As UserSettings)
    Dim hdr As HeaderFooter, ftr As HeaderFooter  ' ҳüҳ�Ŷ���
    Dim shp As Shape                               ' Logoͼ�ζ���
    Dim leftTextBox As Shape, rightTextBox As Shape ' ҳü�����ı���
    Dim frameWidth As Single                       ' ҳü�������
    Dim ҳü���߶��߾� As Single                   ' ҳü���ߵĴ�ֱλ��
    Dim �ı���߶� As Single                       ' ҳü�ı���߶�
    
    ' 1. ҳ�߾����ã����ֺ����棩
    With sec.PageSetup
        If .Orientation = wdOrientPortrait Then  ' ����
            .TopMargin = CentimetersToPoints(settings.topM)
            .BottomMargin = CentimetersToPoints(settings.bottomM)
            .LeftMargin = CentimetersToPoints(settings.leftM)
            .RightMargin = CentimetersToPoints(settings.rightM)
        Else  ' ���
            .TopMargin = CentimetersToPoints(settings.topM_L)
            .BottomMargin = CentimetersToPoints(settings.bottomM_L)
            .LeftMargin = CentimetersToPoints(settings.leftM_L)
            .RightMargin = CentimetersToPoints(settings.rightM_L)
        End If
        .HeaderDistance = CentimetersToPoints(settings.headerDist)
        .FooterDistance = CentimetersToPoints(settings.footerDist)
    End With
    
    ' 2. ҳü������������
    ҳü���߶��߾� = CentimetersToPoints(1.5)  ' ҳü���ߵĴ�ֱλ�ã��̶�1.5cm��
    �ı���߶� = CentimetersToPoints(1)         ' �ı���߶ȹ̶�1cm
    frameWidth = sec.PageSetup.PageWidth - sec.PageSetup.LeftMargin - sec.PageSetup.RightMargin  ' ���ÿ�ȣ�ҳ���-���ұ߾ࣩ
    
    ' 3. ҳü��ʼ�������ԭ�����ݣ�
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    hdr.Range.text = ""
    
    ' 4. Ӧ��ҳü��ʽ��ȷ���ı���Ĭ��ʹ�ø���ʽ��
    Call ����ҳü��ʽ  ' ����/������ʽ
    Call Ӧ��ҳü��ʽ��ҳü(sec)  ' Ӧ�õ���ǰ��ҳü
    
    ' 5. ���ҳüĬ�ϱ߿򣨱�����ţ�
    With hdr.Range.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    End With
    
    ' 6. ��������ı��򣨹���ҳü��ʽ������ʱֱ��Ӧ�ã�
    Set leftTextBox = hdr.Shapes.AddTextBox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=0, Top:=0, width:=frameWidth * 0.45, Height:=�ı���߶�)  ' ռ45%���
    With leftTextBox
        ' ��λ�뻷������
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .WrapFormat.Type = wdWrapNone  ' �������ı�
        .Left = sec.PageSetup.LeftMargin  ' �����ҳ�߾�
        .Top = ҳü���߶��߾� - �ı���߶� * 0.5  ' ��ֱ�����ڵ����Ϸ�
        
        ' �ı����ʽ���ޱ߿򣬱߾�Ϊ0��
        .line.Visible = msoFalse
        .TextFrame.VerticalAnchor = msoAnchorBottom  ' �ı���ֱ����
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0
        .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        .TextFrame.AutoSize = False
    End With
    ' �������ı���Ӧ����ʽ��ֱ�ӹ���HeaderStyle���༭ʱ�Զ��̳У�
    With leftTextBox.TextFrame.TextRange
        .text = settings.headerLeft
        .Style = ActiveDocument.Styles("HeaderStyle")  ' ǿ��Ӧ��ҳü��ʽ
        .ParagraphFormat.alignment = wdAlignParagraphLeft  ' �����
    End With
    
    ' 7. �����Ҳ��ı���ͬ�ϣ�������ʽ��
    Set rightTextBox = hdr.Shapes.AddTextBox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=0, Top:=0, width:=frameWidth * 0.4, Height:=�ı���߶�)  ' ռ40%���
    With rightTextBox
        ' ��λ�뻷������
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .WrapFormat.Type = wdWrapNone
        .Left = sec.PageSetup.LeftMargin + frameWidth * 0.6  ' ����60%��ȣ����ң�
        .Top = ҳü���߶��߾� - �ı���߶� * 0.5  ' ������ı���ֱ����
        
        ' �ı����ʽ��ͬ��ࣩ
        .line.Visible = msoFalse
        .TextFrame.VerticalAnchor = msoAnchorBottom
        .TextFrame.MarginLeft = 0: .TextFrame.MarginRight = 0
        .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        .TextFrame.AutoSize = False
    End With
    ' ����Ҳ��ı���Ӧ����ʽ
    With rightTextBox.TextFrame.TextRange
        .text = settings.headerRight
        .Style = ActiveDocument.Styles("HeaderStyle")  ' ǿ��Ӧ��ҳü��ʽ
        .ParagraphFormat.alignment = wdAlignParagraphRight  ' �Ҷ���
    End With
    
    ' 8. ҳü�������ã���ɫ��ʵ�ߣ�
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
    
    ' 9. Logo�������������Ϸ������½Ƕ�λ��ҳ�߾ࣩ
    If settings.logoPath <> "" And Dir(settings.logoPath) <> "" Then
        Set shp = hdr.Shapes.AddPicture( _
            FileName:=settings.logoPath, _
            LinkToFile:=False, SaveWithDocument:=True)
        With shp
            .LockAspectRatio = msoTrue  ' ���ֱ���
            .Height = CentimetersToPoints(1)  ' �̶��߶�1cm
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .WrapFormat.Type = wdWrapFront  ' ���������Ϸ�
            ' ���½Ƕ�λ������߾࣬���߾ࣩ
            .Left = sec.PageSetup.LeftMargin - .width
            .Top = ҳü���߶��߾� - �ı���߶� * 0.5  ' ���ı���ֱ����
            .ZOrder msoBringToFront  ' ���ڶ���
        End With
    End If
    
    ' 10. ����ҳ��
    Call SetFooter(sec)
End Sub

' === ����ҳ�ţ�����ҳ���ʽ�ͱ�Ź��� ===
Private Sub SetFooter(sec As Section)
    Dim ftr As HeaderFooter
    Dim i As Integer
    Dim firstTitleSection As Section
    Dim isHeading1Found As Boolean
    
    ' 1. ��ʼ��ҳ�ţ�������ݣ�
    Set ftr = sec.Footers(wdHeaderFooterPrimary)
    ftr.Range.text = ""
    isHeading1Found = False  ' ����Ƿ��ҵ�һ������
    
    ' 2. ���ҵ�һ��һ���������ڽڣ��������±�ţ�
    For i = 1 To ActiveDocument.Sections.Count
        If ActiveDocument.Sections(i).Range.Paragraphs(1).Style = "Heading 1" Then
            Set firstTitleSection = ActiveDocument.Sections(i)
            isHeading1Found = True
            Exit For
        End If
    Next i
    
    ' 3. ����ҳ���Ź���
    If isHeading1Found Then
        ' �ӵ�һ��һ���������ڽڿ�ʼ���±�ţ���ʼΪ1��
        With firstTitleSection.Footers(wdHeaderFooterPrimary).PageNumbers
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
    Else
        ' ��һ������ʱ�ӵ�1�ڿ�ʼ���
        With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers
            .RestartNumberingAtSection = True
            .StartingNumber = 1
        End With
    End If
    
    ' 4. ����ҳ���ʽ�����У�Times New Roman��10.5�ţ�
    With ftr.Range.Paragraphs(1).Range
        .Font.name = "Times New Roman"
        .Font.Size = 10.5
        .Font.Color = wdColorBlack
        .Fields.Add Range:=.Characters.Last, Type:=wdFieldPage  ' ����ҳ����
        .ParagraphFormat.alignment = wdAlignParagraphCenter
    End With
End Sub


'���ģ��û������¼�
' === ��ť��Ӧ�õ����� ===
Private Sub cmdApplySection_Click()
    Dim settings As UserSettings
    ' �ռ������ò�Ӧ�õ���ǰѡ�н�
    Call CollectSettings(settings)
    Call ApplySettingsToSection(Selection.Sections(1), settings)
    MsgBox "��Ӧ�õ����ڣ�", vbInformation
End Sub

' === ��ť��ȫ��Ӧ�� ===
Private Sub cmdApplyAll_Click()
    Dim settings As UserSettings
    Dim secCount As Long, i As Long
    Dim ��ʼ�� As Long
    Dim pf As progressForm  ' ���ȴ��壨�����Ѷ��壩
    Dim rsp As VbMsgBoxResult
    Dim rng As Range
    
    ' 1. �ռ�������
    Call CollectSettings(settings)
    
    ' 2. ��ʼ���������ܽ�����������ʼ�ڣ�
    secCount = ActiveDocument.Sections.Count
    ��ʼ�� = ��ȡ������ʼ������()
    
    ' 3. ��ʾ���ȴ���
    ��ʽ_��׼��ҳ������.Hide
    Set pf = New progressForm
    pf.Show vbModeless
    pf.InitForPageSetting secCount, ��ʼ��
    
    ' 4. ȷ���û�����
    rsp = MsgBox("�Ƿ�ʼȫ��ҳüҳ�Ÿ�ʽ����", vbQuestion + vbOKCancel, "ȷ�Ͽ�ʼ")
    If rsp <> vbOK Then
        pf.UpdateProgressBar 0, "�û�ȡ����δ��ʼִ�С�"
        Exit Sub
    End If
    
    ' 5. ѭ��Ӧ�õ����н�
    For i = 1 To secCount
        ' ����Ƿ���Ҫֹͣ
        If pf.stopFlag Then
            pf.UpdateProgressBar pf.FrameProgress.width, "������ֹͣ�������˳�..."
            Exit For
        End If
        
        ' ��������ǰ���ף��������ӻ����飩
        Set rng = ActiveDocument.Sections(i).Range
        rng.Collapse wdCollapseStart
        rng.Select
        On Error Resume Next
        ActiveWindow.ScrollIntoView rng, True
        On Error GoTo 0
        
        ' Ӧ�����õ���ǰ��
        Call ApplySettingsToSection(ActiveDocument.Sections(i), settings)
        
        ' ���½���
        pf.UpdateProgressBar CLng(200# * i / secCount), _
            "�Ѿ���ɵ� " & i & " ��ҳüҳ�����ã��� " & secCount & " ��"
        DoEvents
    Next i
    
    ' 6. �����ʾ
    If Not pf.stopFlag Then
        pf.UpdateProgressBar 200, "�Ѿ����ȫ��ҳüҳ�����ã�������ȡ�����˳���"
    End If

'    ��ʽ_��׼��ҳ������.Show
    
End Sub


' === ��ť��ȡ�� ===
Private Sub cmdCancel_Click()
    Unload Me  ' �رձ�
End Sub


Private Sub txtHeaderLeft_Change()

End Sub

' === ����ʼ�������ÿؼ�Ĭ��ֵ ===
Private Sub UserForm_Initialize()
    ' ����߾�Ĭ��ֵ
    txtTop.text = "2.5": txtBottom.text = "2.5": txtLeft.text = "3": txtRight.text = "3"
    ' ���߾�Ĭ��ֵ
    txtTopL.text = "3": txtBottomL.text = "3": txtLeftL.text = "2.5": txtRightL.text = "2.5"
    ' ҳü�ı�Ĭ��ֵ���������У�
    txtHeaderLeft.text = "������������MHP0-1403��Ԫ" & Chr(13) & "73-04�ؿ�����(��Ǩ)����ס����Ŀ"
    txtHeaderRight.text = "ʩ����֯���"
    ' ҳüҳ�ž���Ĭ��ֵ
    txtHeaderDist.text = "1.5": txtFooterDist.text = "1.75"
    ' LogoĬ��·����ʾ����
    txtLogo.text = "C:\Users\Tony Zhang\Desktop\logo.png"
End Sub

' === ���Logo�����ļ�ѡ����ѡ��ͼƬ ===
Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "ѡ��LogoͼƬ"
        .Filters.Clear
        .Filters.Add "ͼƬ�ļ�", "*.png; *.jpg; *.jpeg; *.bmp; *.gif"
        If .Show = -1 Then Me.txtLogo.text = .SelectedItems(1)  ' ѡ�к����·��
    End With
End Sub
'==================== �ɴ��壺�н���׼��ҳ������ ====================
' ˵����
' ��һ���� 4 ������������䴰�塰Զ�̵��á��ɴ����߼����Ķ���С��
' ������host �������䴰�壨Me���������������Ȱ� host ��ֵͬ�������������ִ�У�
' �������ڲ�ֱ�ӵ������еİ�ť�¼���*_Click����������ɴ��롣

' ������0���ѹ�������Ŀؼ�ֵͬ���������壨˽��С���ߣ�����
Private Sub PS_CopyFromHost(ByVal host As Object)
    On Error Resume Next
    ' ����
    Me.txtTop.text = host.txtTop.text
    Me.txtBottom.text = host.txtBottom.text
    Me.txtLeft.text = host.txtLeft.text
    Me.txtRight.text = host.txtRight.text
    ' ���
    Me.txtTopL.text = host.txtTopL.text
    Me.txtBottomL.text = host.txtBottomL.text
    Me.txtLeftL.text = host.txtLeftL.text
    Me.txtRightL.text = host.txtRightL.text
    ' ҳü/ҳ��/LOGO
    Me.txtHeaderLeft.text = host.txtHeaderLeft.text
    Me.txtHeaderRight.text = host.txtHeaderRight.text
    Me.txtLogo.text = host.txtLogo.text
    Me.txtHeaderDist.text = host.txtHeaderDist.text
    Me.txtFooterDist.text = host.txtFooterDist.text
    On Error GoTo 0
End Sub

'��һ��������ڣ���ʼ������ host ��ֵͬ����������ִ��ҵ��
Public Sub PS_InitFromHost(ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
End Sub

'������������ڣ�Ӧ�õ����ڣ���ѡͬ�� host �� �����壬����������а�ť�߼���
Public Sub PS_ApplySection(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    ' ֱ�ӵ��������еġ�Ӧ�õ����ڡ���ť�¼�����
    On Error Resume Next
    Call cmdApplySection_Click
    On Error GoTo 0
End Sub

'������������ڣ�ȫ��Ӧ�ã���ѡͬ�� host �� �����壬����������а�ť�߼���
Public Sub PS_ApplyAll(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    On Error Resume Next
    Call cmdApplyAll_Click
    On Error GoTo 0
End Sub

'���ģ�������ڣ���� Logo������ host ��ͬ�����ٵ��������е�ѡ���߼���
Public Sub PS_BrowseLogo(Optional ByVal host As Object)
    If Not host Is Nothing Then PS_CopyFromHost host
    On Error Resume Next
    Call cmdBrowse_Click
    On Error GoTo 0
End Sub
'==================== /�ɴ��壺�н���׼��ҳ������ ====================


