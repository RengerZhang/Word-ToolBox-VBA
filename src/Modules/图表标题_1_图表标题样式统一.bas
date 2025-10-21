Attribute VB_Name = "ͼ�����_1_ͼ�������ʽͳһ"
Option Explicit

'==============================================
' ģ�飺ͼƬ������ƥ�䣨�����ȣ�
'==============================================

'��һ����ڣ�ͼƬ������ʽͳһ�������ȣ�
'��Ҫ��ȫ��ÿ��ͼƬ�·��ĵ�һ���ǿն������óɡ�ͼƬ���⡱��ʽ
Public Sub ͼƬ������ʽͳһ_������()
    Dim doc As Document: Set doc = ActiveDocument
    
    '������׼�����ȴ���Ŀ����ʽ
    Dim pf As progressForm
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "ͼƬ����ƥ�䣨�����ȣ�"
        pf.FrameProgress.width = 0
        pf.LabelPercentage.caption = "0%"
        pf.TextBoxStatus.text = "��ʼƥ��ͼƬ���⡭��"
        pf.Show vbModeless
        DoEvents
    End If
    
    EnsureParaStyleExists doc, "ͼƬ����"  ' ���������򴴽����̳й�ϵ������ʽ�������裬����ֻ��֤�У�
    
    '����������������InlineShapes + ����Shapes�е�ͼƬ��
    Dim total As Long, cntInline As Long, cntShape As Long
    cntInline = doc.InlineShapes.Count
    cntShape = CountPictureShapes(doc)
    total = cntInline + cntShape
    
    If total = 0 Then
        UpdateBar pf, 200, 200, "δ�����κ�ͼƬ��"
        GoTo CLEANUP
    End If
    
    '���ģ�������� InlineShapes��Ƕ��ʽͼƬ��
    Dim i As Long, ils As InlineShape
    i = 0
    For Each ils In doc.InlineShapes
        i = i + 1
        If Not CaptionForInlineShape(doc, ils, "ͼƬ����") Then
            StatusPulse pf, "���������� " & i & " �� Inline ͼƬδ�ҵ��·��ǿն��䡣"
        End If
        
        UpdateBar pf, CInt(200# * i / total), 200, "����ͼƬ��Inline����" & i & "/" & total
        If Not pf Is Nothing Then If pf.stopFlag Then GoTo CLEANUP
    Next
    
    '���壩��������� Shapes������ʽͼƬ��
    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape(s) Then
            i = i + 1
            If Not CaptionForShape(doc, s, "ͼƬ����") Then
                StatusPulse pf, "���������� " & i & " �Ÿ���ͼƬδ�ҵ��·��ǿն��䡣"
            End If
            
            UpdateBar pf, CInt(200# * i / total), 200, "����ͼƬ����������" & i & "/" & total
            If Not pf Is Nothing Then If pf.stopFlag Then GoTo CLEANUP
        End If
    Next
    
    StatusPulse pf, "ͼƬ����ƥ����ɣ������� " & total & " ��ͼƬ��"

CLEANUP:
    If Not pf Is Nothing Then Unload pf
End Sub


'��������ڣ��������ʽͳһ�������ȣ�
' �߼���ԭ�������ƥ�䡱һ�£���ÿ��������Ѱ�ҵ�һ���ǿնβ���Ϊ�������⡿��ʽ
Public Sub �������ʽͳһ_������()
    Dim doc As Document: Set doc = ActiveDocument
    
    Dim pf As progressForm
    On Error Resume Next
    Set pf = VBA.UserForms.Add("ProgressForm")
    On Error GoTo 0
    If Not pf Is Nothing Then
        pf.caption = "�����ƥ�䣨�����ȣ�"
        pf.FrameProgress.width = 0
        pf.LabelPercentage.caption = "0%"
        pf.TextBoxStatus.text = "��ʼƥ�����⡭��"
        pf.Show vbModeless
        DoEvents
    End If
    
    EnsureParaStyleExists doc, "������"
    
    Dim n As Long: n = doc.Tables.Count
    If n = 0 Then
        UpdateBar pf, 200, 200, "�ĵ���û�б��"
        GoTo EXIT_B
    End If
    
    Dim i As Long, tbl As Table, rng As Range, prevPara As Paragraph
    For i = 1 To n
        Set tbl = doc.Tables(i)
        Set rng = tbl.Range: rng.Collapse wdCollapseStart
        Set prevPara = rng.Paragraphs(1).Previous
        
        Do While Not prevPara Is Nothing
            Dim t As String: t = ��������ı�(prevPara.Range.text)
            If Len(t) > 0 Then
                prevPara.Style = doc.Styles("������")
                Exit Do
            End If
            Set prevPara = prevPara.Previous
        Loop
        
        UpdateBar pf, CInt(200# * i / n), 200, "������" & i & "/" & n
        If Not pf Is Nothing Then If pf.stopFlag Then Exit For
    Next
    
    StatusPulse pf, "�����ƥ����ɣ������� " & n & " �ű�"

EXIT_B:
    If Not pf Is Nothing Then Unload pf
End Sub


'======================== ���������� ========================

'���ߣ��� InlineShape ���á��·���һ���ǿնΡ�ΪĿ����ʽ
Private Function CaptionForInlineShape(ByVal doc As Document, ByVal ils As InlineShape, ByVal styleName As String) As Boolean
    On Error GoTo SAFE_EXIT
    Dim p As Paragraph
    Set p = NextNonEmptyPara(ils.Range.Paragraphs(1))
    If Not p Is Nothing Then
        p.Style = doc.Styles(styleName)
        CaptionForInlineShape = True
    End If
SAFE_EXIT:
End Function

'���ˣ��Ը��� Shape ���á�ê���·���һ���ǿնΡ�ΪĿ����ʽ
Private Function CaptionForShape(ByVal doc As Document, ByVal s As Shape, ByVal styleName As String) As Boolean
    On Error GoTo SAFE_EXIT
    If s Is Nothing Then Exit Function
    Dim anchorPara As Paragraph
    Set anchorPara = s.anchor.Paragraphs(1)
    Dim target As Paragraph
    Set target = NextNonEmptyPara(anchorPara)
    If Not target Is Nothing Then
        target.Style = doc.Styles(styleName)
        CaptionForShape = True
    End If
SAFE_EXIT:
End Function

'���ţ��Ӹ������俪ʼ�����¡�Ѱ�ҵ�һ���ǿնΣ�������ǰ�Σ�
Private Function NextNonEmptyPara(ByVal p As Paragraph) As Paragraph
    Dim q As Paragraph
    If p Is Nothing Then Exit Function
    Set q = p.Next
    Do While Not q Is Nothing
        If Len(��������ı�(q.Range.text)) > 0 Then
            Set NextNonEmptyPara = q
            Exit Function
        End If
        Set q = q.Next
    Loop
End Function

'��ʮ���ж� Shape �Ƿ�ΪͼƬ��msoPicture �� msoLinkedPicture��
Private Function IsPictureShape(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsPictureShape = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

'��ʮһ��ͳ���ĵ��еġ�ͼƬ�͡����� Shape ����
Private Function CountPictureShapes(ByVal doc As Document) As Long
    Dim s As Shape, n As Long
    For Each s In doc.Shapes
        If IsPictureShape(s) Then n = n + 1
    Next
    CountPictureShapes = n
End Function

'��ʮ������֤������ʽ���ڣ��������򴴽���
Private Sub EnsureParaStyleExists(ByVal doc As Document, ByVal styleName As String)
    On Error Resume Next
    Dim st As Style
    Set st = doc.Styles(styleName)
    If st Is Nothing Then
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0
End Sub

'��ʮ�������ߣ��������ɼ��ı���ȥβ���/ȫ�ǿո� Trim��
Private Function ��������ı�(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")                 ' ��Ԫ�������
    s = Replace$(s, ChrW(&H3000), " ")          ' ȫ�ǿո�ת���
    ��������ı� = Trim$(s)
End Function

'��ʮ�ģ����ȸ�����״̬������������㣩
Private Sub StatusPulse(ByVal pf As progressForm, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.TextBoxStatus.text = pf.TextBoxStatus.text & vbCrLf & msg
        pf.TextBoxStatus.SelStart = Len(pf.TextBoxStatus.text)
        pf.TextBoxStatus.SelLength = 0
        pf.Repaint
    End If
    DoEvents
    On Error GoTo 0
End Sub

'��ʮ�壩���ȸ��������½�������ProgressForm �Ľ������ܿ� 200px��
Private Sub UpdateBar(ByVal pf As progressForm, ByVal cur As Long, ByVal total As Long, ByVal msg As String)
    On Error Resume Next
    If Not pf Is Nothing Then
        pf.UpdateProgressBar cur, msg         ' ���� 0~200 �Ŀ��
    End If
    DoEvents
    On Error GoTo 0
End Sub


