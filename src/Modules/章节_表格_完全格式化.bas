Attribute VB_Name = "�½�_���_��ȫ��ʽ��"
Sub �½ڱ���ʽ������()
    Dim oRow As row
    Dim oCell As cell
    Dim tb As Table
    Dim myParagraph As Paragraph, n As Integer
    Dim progressForm As progressForm ' ����֮ǰ������ ProgressForm
    Dim ��ǰ�½� As Range
    Dim ������ As Integer
    Dim ��ǰ�� As Integer
    
    ' ��һ�����á���׼�������ʽ������
    EnsureStandardTableStyle
    
    
    ' (��)��ȡ�½ڱ�������ͱ��
    ' ��ȡ������ڵĽ�
    ��ǰ�� = Selection.Sections(1).Index
    
    ' ��ȡ��ǰ�ڵķ�Χ
    Set ��ǰ�½� = ActiveDocument.Sections(��ǰ��).Range
    
    ' ����ý��еı������
    ������ = ��ǰ�½�.Tables.Count
    i = ������
    
    ' ��������ʾ���ȴ���
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' ��ģ̬���壬����������Լ���ִ��
    progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' ��ʼ�����ȴ���
    progressForm.TextBoxStatus.text = "���ڹ��� " & i & " ��������ڿ�ʼ��ʽ��..."

    
    For r = 1 To i
        If progressForm.stopFlag Then
            ' ��������ǿ��ֹͣ��ť���˳�ѭ��
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ֹͣ�������˳�..."
            Exit For
        End If
        
        tbl = ��ǰ�½�.Tables(r)
        Set tb = ��ǰ�½�.Tables(r)
        
        tbl.Select
        Selection.Style = "��׼�������ʽ"
        '    ���ñ���������ú���
        Call �����������(ActiveDocument.Tables(r))
        
        
        
        '  ������Ԫ�������ڿ���,Ϊ�˷�ֹ�ϲ���Ԫ������
        For Each oCell In tbl.Cells
            oCell.Select
            With Selection
                .Borders.OutsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineWidth = wdLineWidth050pt
            End With
            
            Selection.SelectRow
            Selection.rows.AllowBreakAcrossPages = enable
            
            n = 1
            For Each myParagraph In Selection.Paragraphs
                If Len(Trim(myParagraph.Range)) = 1 Then
                    myParagraph.Range.Delete
                    n = n + 1
                End If
            Next
            
        Next oCell

        '  ѡ�б�����������
        With tbl.Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth150pt ' ���1.5��
            .OutsideColor = wdColorBlack
        End With
        
        '  �������мӴֲ���ÿҳ�ظ�
        tbl.Select
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
    ' ��ɺ��� TextBox ����ʾ�����Ϣ
    
    ' ׷�Ӽ�¼�� TextBox
    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "����ʽ������ϣ�"
    'progressForm.Hide
    
    Exit Sub
    
    ' ��ϸ������ʾ��Ϣ
    MsgBox "����ʽ������ɣ���ִ�����²�����" & vbCrLf & _
        "1. Ӧ��""���""��ʽ" & vbCrLf & _
        "2. �Զ���Ӧ���ڿ��" & vbCrLf & _
        "3. ���õ�Ԫ�����±߾�Ϊ0" & vbCrLf & _
        "4. �о��ж��룬ȡ�����ֻ���" & vbCrLf & _
        "5. �������ã�����(����)��Times New Roman(Ӣ��)��10.5����" & vbCrLf & _
        "6. ��Ԫ�����ݴ�ֱ��ˮƽ����" & vbCrLf & _
        "7. �������ã������о࣬��ǰ���࣬������" & vbCrLf & _
        "8. �����ɫ���и���Ϊ��Сֵ0.6cm" & vbCrLf & _
        "9. �����ڿ���(0.5��)�������(1.5��)" & vbCrLf & _
        "10. ��ֹ����п�ҳ����" & vbCrLf & _
        "11. ɾ����Ԫ���ڿն���" & vbCrLf & _
        "12. ���мӴֲ�����Ϊÿҳ�ظ�������", _
        vbInformation, "�������"
    
End Sub
