Attribute VB_Name = "ȫ��_���_ͳһ������ʽ"
Sub ȫ�ı����ʽ��ʽ��()
    Dim tb As Table
    Dim myParagraph As Paragraph, N As Integer
    Dim progressForm As progressForm ' ����֮ǰ������ ProgressForm
    
    ' ���á���׼�������ʽ������
    EnsureStandardTableStyle
    
    ' ����ȫ�ı������
    i = ActiveDocument.Tables.Count
    
    ' ��������ʾ���ȴ���
    Set progressForm = New progressForm
    progressForm.Show vbModeless ' ��ģ̬���壬����������Լ���ִ��
    progressForm.TextBoxStatus.text = "ȫ�Ĺ���" & i & "��������ڿ�ʼ��ʽ�������ʽ..."
     progressForm.UpdateProgressBar 0, "Processing table 1 of " & i   ' ��ʼ�����ȴ���
    
    
    For r = 1 To i
        If progressForm.stopFlag Then
            progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "������ֹͣ�������˳�..."
            Exit For
        End If
        
        tbl = ActiveDocument.Tables(r)
        Set tb = ActiveDocument.Tables(r)
        
        tbl.Select
        Selection.Style = "��׼�������ʽ"
       
        ' ���½�������״̬�ı���
        progressForm.UpdateProgressBar (r / i) * 200, "Processing table " & r & " of " & i
        ' ȷ���������
        DoEvents
        
    Next r

    progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & vbCrLf & "�����ʽ������ϣ�"
    
    Exit Sub
    
End Sub
