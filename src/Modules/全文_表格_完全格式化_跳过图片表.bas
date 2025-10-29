Attribute VB_Name = "ȫ��_���_��ȫ��ʽ��_����ͼƬ��"
' =======����������������ʽ��������ģ�鶥���������⣩=======
Const S_TABLE_PIC As String = "ͼƬ��λ��"     ' �����ʽ��ͼƬ��
Const S_TABLE_NOR As String = "��׼�����ʽ"   ' �����ʽ��һ���
Const S_PARA_IMG As String = "ͼƬ��ʽ"        ' ������ʽ�����ڡ���ͼ��Ԫ��
Const S_PARA_CAP As String = "ͼƬ����"        ' ������ʽ�����ڡ���ͼ��Ԫ�񡱣����⣩


'====================��A��������������ֱ��ִ�� ====================
Public Sub ȫ�ı���ʽ��_������1( _
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

' ֻ�ڡ�������ĵط���һ��ͼƬ���֧��һ���ԭ�߼�����
Private Sub ȫ�ı���ʽ��_����( _
    ByVal optThick As Boolean, _
    ByVal optHeadBold As Boolean, _
    ByVal optFontPt As Single, _
    ByVal optFontName As String _
)
    Dim tb As Table
    Dim oCell As cell
    Dim myParagraph As Paragraph
    Dim N As Integer
    Dim progressForm As progressForm
    Dim i As Long, r As Long

    ' ��ʽ׼������ԭ�еģ�
    EnsureStandardTableStyle
    ActiveDocument.Styles("��׼�������ʽ").Font.Size = optFontPt

    ' ���ȴ��壨��ԭ�еģ�
    i = ActiveDocument.Tables.Count
    Set progressForm = New progressForm
    progressForm.Show vbModeless
    progressForm.TextBoxStatus.text = "ȫ�Ĺ��� " & i & " ��������ڿ�ʼ��ʽ��..."
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i

    ' ===== ����� =====
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ֹͣ�������˳�..."
            Exit For
        End If

        Set tb = ActiveDocument.Tables(r)

        ' =========��������ͼƬ���ж���nͼ �� ��ֵ n+1��=========
        Dim nInline As Long, nShape As Long, nImg As Long
        Dim totalCells As Long, imgCells As Long, txtEst As Long
        Dim isPic As Boolean

        nInline = tb.Range.InlineShapes.Count
        nShape = SafeShapeCount_InRange(tb.Range)   ' ����ͼ��������ȫ��
        nImg = nInline + nShape

        totalCells = tb.Range.Cells.Count           ' ���ݺϲ���Ԫ��
        imgCells = CountImageCells(tb)              ' ��ͼ��Ԫ����
        txtEst = totalCells - imgCells              ' �������ֵ�Ԫ��
        isPic = (nImg > 0 And txtEst <= nImg + 1)

        If isPic Then
            ' =======��������ͼƬ����=======
            Call ��ʽ��_ͼƬ��λ��(tb)            ' ֻ����Ҫ���������
            ' ���� & ��һ�ű�
            progressForm.UpdateProgressBar CLng((r / i) * 200), _
                "Processing table " & r & " of " & i & "��ͼƬ��λ��"
            DoEvents
            GoTo NextTable_Continue                 ' ����һ����ԭ����
        End If
        ' =========����������֧������������롰ԭ�е�һ������̡�=========

        ' ========�����±������ԭ���벻����========
        ' 1) Ӧ����ʽ + ��������
        tb.Select: Selection.Style = "��׼�������ʽ"
        ����������� tb                                   ' �� ��ԭ���Ĺ�����
        ' 2) �ڿ��߹̶� 0.5 ��
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With
            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = False
            N = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    N = N + 1
                End If
            Next
        Next oCell

        ' 3) ����ߣ������� 1.5 / 0.5
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = IIf(optThick, wdLineWidth150pt, wdLineWidth050pt)
            .OutsideColor = wdColorBlack
        End With

        ' 4) ���мӴ� + ÿҳ�ظ�
        tbl.Select
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = optHeadBold
        Selection.rows.HeadingFormat = True

        ' 5) ����
        progressForm.UpdateProgressBar CLng((r / i) * 200), _
            "Processing table " & r & " of " & i
        DoEvents

NextTable_Continue:
    Next r

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "����ʽ������ϣ�"
End Sub


'====================��C�����ݱ���������ڣ�����ȡֵ���ٵ����ģ� ====================
Public Sub ȫ�ı���ʽ������1()
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

'������ȫͳ��ĳ��Χ�ڸ���ͼ������ͼ������
Private Function SafeShapeCount_InRange(ByVal rng As Range) As Long
    On Error Resume Next
    SafeShapeCount_InRange = rng.ShapeRange.Count
    On Error GoTo 0
End Function

'������Ԫ�����Ƿ���ͼƬ�����ڻ򸡶���
Private Function CellHasImage(ByVal c As cell) As Boolean
    CellHasImage = (c.Range.InlineShapes.Count > 0) Or (SafeShapeCount_InRange(c.Range) > 0)
End Function

'����ͳ�ơ���ͼ��Ԫ������
Private Function CountImageCells(ByVal tb As Table) As Long
    Dim c As cell, N As Long
    For Each c In tb.Range.Cells
        If CellHasImage(c) Then N = N + 1
    Next c
    CountImageCells = N
End Function

'����ͼƬ��ר�ã��������Ӧ����ͼ��Ԫ�����ͼƬ��ʽ�����������ͼƬ���⡿
Private Sub ��ʽ��_ͼƬ��λ��(ByVal tb As Table)
    ' ���������Ӧ
    tb.AutoFitBehavior wdAutoFitWindow

    ' ��ѡ��������ϡ�ͼƬ��λ���ı����ʽ�����ھ��ã�������Ҳ������
    On Error Resume Next
    tb.Style = S_TABLE_PIC
    On Error GoTo 0

    ' ��Ԫ�����ö�����ʽ
    Dim c As cell
    For Each c In tb.Range.Cells
        If CellHasImage(c) Then
            On Error Resume Next
            c.Range.Style = ActiveDocument.Styles(S_PARA_IMG)   ' ͼƬ��ʽ
            On Error GoTo 0
        Else
            On Error Resume Next
            c.Range.Style = ActiveDocument.Styles(S_PARA_CAP)   ' ͼƬ����
            On Error GoTo 0
        End If
    Next c
End Sub





