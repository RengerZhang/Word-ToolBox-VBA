Attribute VB_Name = "ȫ��_���_����Ӵ�"
Sub ȫ�ı�����мӴ�()
    Dim tb As Table
    Dim myParagraph As Paragraph, n As Integer
    Dim progressForm As progressForm ' ����֮ǰ������ ProgressForm
    
    
    ' ����ȫ�ı������
    i = ActiveDocument.Tables.Count
    
    ' ��������ʾ���ȴ���
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' ��ģ̬���壬����������Լ���ִ��
    progressForm.TextBoxStatus.text = "ȫ�Ĺ���" & i & "��������ڿ�ʼ�Ӵֱ���..."
     progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' ��ʼ�����ȴ���
    
    
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ֹͣ�������˳�..."
            Exit For
        End If
        
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)

        '  �������мӴֲ���ÿҳ�ظ�
        tbl.rows.Select
        Selection.rows.HeadingFormat = wdUndefined
        Selection.rows.HeadingFormat = False
        tbl.Cells(1).Select
        Selection.SelectRow
        Selection.Range.bold = True
        Selection.rows.HeadingFormat = True
        ' ���½�������״̬�ı���
        progressForm.UpdateProgressBar (r / i) * 200, "Processing table " & r & " of " & i
        ' ȷ���������
        DoEvents
        
    Next r

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������Ӵ���ϣ�"
    
    Exit Sub
    
End Sub

'==========================================================
' ��ǰ����ʽ���ƣ�������ڵ���һ�ű�
' ������
'   thickOuter   ���Ӵ֣�True=1.5����False=0.5����
'   firstRowBold ���мӴ�
'   headerRepeat ����ÿҳ�ظ�
'   allowBreak   �����ҳ���У�����
'   fontSizePt   �����ֺţ�����
'==========================================================
Public Sub ��ǰ����ʽ����(thickOuter As Boolean, _
                           firstRowBold As Boolean, _
                           headerRepeat As Boolean, _
                           allowBreak As Boolean, _
                           fontSizePt As Single)
    Dim tb As Table
    '��һ���õ�������ڱ��
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "���δ�ڱ���С�", vbExclamation
        Exit Sub
    End If
    Set tb = Selection.Tables(1)

    '�����������ֺţ����㡰ֻ��һ��ֵ����˼·��
    tb.Range.Font.Size = fontSizePt

    '�������ڿ��ߣ��̶� 0.5 ����һ��д����
    tb.Borders.InsideLineStyle = wdLineStyleSingle
    tb.Borders.InsideLineWidth = wdLineWidth050pt

    '���ģ�����ߣ���=1.5 ������=0.5 ����ֱ�ӿ����
    With tb.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideLineWidth = IIf(thickOuter, wdLineWidth150pt, wdLineWidth050pt)
        .OutsideColor = wdColorBlack
    End With

    '���壩���мӴ� & �����ظ���ֱ�ӿ���������
    If tb.rows.Count > 0 Then
        tb.rows(1).Range.bold = firstRowBold
        tb.rows(1).HeadingFormat = headerRepeat
    End If

    '�����������Ƿ������ҳ����
    tb.rows.AllowBreakAcrossPages = allowBreak
End Sub


