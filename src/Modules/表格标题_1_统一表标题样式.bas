Attribute VB_Name = "������_1_ͳһ�������ʽ"
Sub �������ʽͳһ()
    Dim doc As Document
    Dim Para As Paragraph
    Dim lvl As Long
    Dim text As String
    Dim tbl As Table
    Dim rng As Range
    
    Set doc = ActiveDocument

      '======== �����⴦���ݴ��������հ׶Σ�========
    For Each tbl In doc.Tables
        Set rng = tbl.Range
        rng.Collapse wdCollapseStart
        
        Dim prevPara As Paragraph
        Set prevPara = rng.Paragraphs(1).Previous
        
        Do While Not prevPara Is Nothing
            Dim t As String
            t = ��������ı�(prevPara.Range.text)
            If Len(t) > 0 Then
                prevPara.Style = doc.Styles("������")
                Exit Do
            End If
            Set prevPara = prevPara.Previous
        Loop
    Next tbl
    
End Sub
'������ ���ߣ��������ɼ��ı���ȥ����β/��Ԫ���ǡ�ȫ�ǿո� Trim��
Private Function ��������ı�(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")                ' ��Ԫ�������
    s = Replace(s, ChrW(&H3000), " ")         ' ȫ�ǿո�ת���
    ��������ı� = Trim$(s)
End Function

