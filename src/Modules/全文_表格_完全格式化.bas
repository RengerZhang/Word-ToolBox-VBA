Attribute VB_Name = "ȫ��_���_��ȫ��ʽ��"
'====================��A��������������ֱ��ִ�� ====================
Public Sub ȫ�ı���ʽ��_������( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    Optional ByVal optFontName As String = "���" _
)
    '��һ������У��
    If optFontPt <= 0! Then
        MsgBox "�ֺŲ�����Ч��", vbExclamation
        Exit Sub
    End If
    '�������������
    ȫ�ı���ʽ��_���� optThick, optHeadBold, optFontPt, optFontName
End Sub


'====================��B�����ģ�ֻ��һ���¡����������ܸ�ʽ�� ====================
Private Sub ȫ�ı���ʽ��_����( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    ByVal optFontName As String _
)
    '��һ������
    Dim tb As Table
    Dim oCell As cell
    Dim myParagraph As Paragraph
    Dim n As Integer
    Dim progressForm As progressForm
    Dim i As Long, r As Long

    '��������ʽ׼�����ֺ����Բ�����
    EnsureStandardTableStyle
    ActiveDocument.Styles("��׼�������ʽ").Font.Size = optFontPt

    '���������� + ���ȴ���
    i = ActiveDocument.Tables.Count
    Set progressForm = New progressForm
    progressForm.Show vbModeless
    progressForm.TextBoxStatus.text = "ȫ�Ĺ��� " & i & " ��������ڿ�ʼ��ʽ��..."
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i
   

    '���ģ������
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ֹͣ�������˳�..."
            Exit For
        End If

        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)

        ' 1) Ӧ����ʽ + ��������
        tbl.Select
        Selection.Style = "��׼�������ʽ"
        ����������� tb

        ' 2) �ڿ��߹̶� 0.5 ��
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With

            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = False  ' ���������ԭ�ȵ� enable ���������滻Ϊ�ñ���

            n = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    n = n + 1
                End If
            Next
        Next oCell

        ' 3) ����ߣ������� 1.5 / 0.5
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = IIf(optThick, wdLineWidth150pt, wdLineWidth050pt)
            .OutsideColor = wdColorBlack
        End With

        ' 4) ���мӴ� + ÿҳ�ظ����Ӵ��ɲ�������
        tbl.Select
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = optHeadBold     ' �� ֻ�������ܿ�
        Selection.rows.HeadingFormat = True

        ' 5) ����
        progressForm.UpdateProgressBar CLng((r / i) * 200), _
            "Processing table " & r & " of " & i
        DoEvents
    Next r

    '���壩�����ʾ
    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "����ʽ������ϣ�"
End Sub


'====================��C�����ݱ���������ڣ�����ȡֵ���ٵ����ģ� ====================
Public Sub ȫ�ı���ʽ������()
    Dim dlg As ��׼����ʽ������
    Set dlg = New ��׼����ʽ������
    dlg.Show vbModeless
    If dlg.Canceled Then
        MsgBox "��ȡ����", vbInformation
        Exit Sub
    End If

    ȫ�ı���ʽ��_���� _
        dlg.SelectedThickOuter, _
        dlg.SelectedFirstRowBold, _
        dlg.SelectedFontSizePt, _
        dlg.SelectedFontSizeName
End Sub


' ������������������ã�֧���ⲿ�����ֺţ�����
Public Sub �����������(tb As Table, Optional ByVal fontPt As Single = 0!)
    '��һ���Զ����� + �ڱ߾�
    tb.AutoFitBehavior (wdAutoFitWindow)
    tb.TopPadding = PixelsToPoints(0, True)
    tb.BottomPadding = PixelsToPoints(0, True)
    tb.LeftPadding = PixelsToPoints(0, True)
    tb.RightPadding = PixelsToPoints(0, True)

    '������׼���ֺţ�δ�������ʽȡ����ȡ��������� 10.5
    Dim pt As Single
    If fontPt > 0! Then
        pt = fontPt
    Else
        On Error Resume Next
        pt = ActiveDocument.Styles("��׼�������ʽ").Font.Size
        On Error GoTo 0
        If pt <= 0! Then pt = 10.5
    End If

    '��������ʽ���������
    tb.Select
    With Selection
        .rows.alignment = wdAlignRowCenter
        .rows.WrapAroundText = False

        ' �����������
        .Font.NameFarEast = ""
        .Font.NameAscii = ""
        .Range.bold = False

        ' �������壨�ֺ��� pt��
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameFarEast = "����"
        .Range.Font.Size = pt

        ' ��Ԫ�����������
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        With .ParagraphFormat
            .CharacterUnitFirstLineIndent = 0
            .alignment = wdAlignParagraphCenter
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
            .LeftIndent = 0
            .RightIndent = 0
        End With

        ' �����ɫ
        .Shading.BackgroundPatternColor = wdColorAutomatic

        ' �и�
        .rows.HeightRule = wdRowHeightAtLeast
        .rows.Height = CentimetersToPoints(0.6)
    End With
End Sub



