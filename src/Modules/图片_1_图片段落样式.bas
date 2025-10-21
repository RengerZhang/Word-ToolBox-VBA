Attribute VB_Name = "ͼƬ_1_ͼƬ������ʽ"
'==============================================================
' ���ܣ�ͳһȫ��ͼƬ�Ķ�����ʽΪ��ͼƬ��ʽ��
' Լ������ʽ�ڡ���ʽ����ģ�顱�ж��壻��ģ��ֻ������
' �ų���ҳüҳ�ŵȷ����� Story�����ų�����е�ͼƬ
' ������ʹ�� ProgressForm ��ʾ���ȣ�����ֹ
'==============================================================
Public Sub ͳһͼƬ������ʽ_ʹ�ý��ȴ���()
    Dim doc As Document: Set doc = ActiveDocument
    Dim styPicPara As Style
    
    '��һ��У��Ŀ����ʽ�Ƿ��ѵ��루��������ֻ���ã�
    On Error Resume Next
    Set styPicPara = doc.Styles("ͼƬ��ʽ")
    On Error GoTo 0
    If styPicPara Is Nothing Then
        MsgBox "δ�ҵ���ʽ��ͼƬ��ʽ����" & vbCrLf & _
               "�����ڡ���ʽ���롿�е������ʽ�������С�", vbExclamation, "ͼƬ������ʽͳһ"
        Exit Sub
    End If
    
    '������ͳ��ͼƬ������ֻ������ Story��
    Dim nInline As Long, nShape As Long, total As Long
    nInline = CountInlinePictures_MainStory(doc)
    nShape = CountFloatingPictures_MainStory(doc)
    total = nInline + nShape
    
    If total = 0 Then
        MsgBox "�ĵ���δ��⵽�κ�ͼƬ�����Ĳ��֣���", vbInformation
        Exit Sub
    End If
    
    '�������򿪽��ȴ���
    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "ͳһͼƬ������ʽ��ͼƬ��ʽ��"
    pf.FrameProgress.width = 0
    pf.LabelPercentage.caption = "0%"
    pf.TextBoxStatus.text = "����⵽ͼƬ��" & total & "��Inline=" & nInline & "��Floating=" & nShape & "��"
    pf.Show vbModeless
    DoEvents
    
    Application.ScreenUpdating = False
    
    Dim done As Long, changed As Long, unchanged As Long
    Dim p As Paragraph
    
    '���ģ����� InlineShapes��Ƕ��ʽ��
    Dim ils As InlineShape
    For Each ils In doc.InlineShapes
        If pf.stopFlag Then GoTo EARLY_OUT
        If IsInlinePicture(ils) Then
            Set p = ils.Range.Paragraphs(1)
            ' ���������� Story
            If p.Range.StoryType = wdMainTextStory Then
                If Not ParaHasStyle(p, styPicPara) Then
                    p.Style = styPicPara
                    changed = changed + 1
                Else
                    unchanged = unchanged + 1
                End If
                done = done + 1
                pf.UpdateProgressBar ProgressPixels(done, total), "����Inline����" & done & "/" & total
            End If
        End If
        DoEvents
    Next
    
    '���壩������ Shapes������ͼƬ��
    Dim s As Shape
    For Each s In doc.Shapes
        If pf.stopFlag Then GoTo EARLY_OUT
        If IsFloatingPicture(s) Then
            ' ȡê�����ڶΣ������ڱ��������ų���
            Set p = s.anchor.Paragraphs(1)
            If p.Range.StoryType = wdMainTextStory Then
                If Not ParaHasStyle(p, styPicPara) Then
                    p.Style = styPicPara
                    changed = changed + 1
                Else
                    unchanged = unchanged + 1
                End If
                done = done + 1
                pf.UpdateProgressBar ProgressPixels(done, total), "������������" & done & "/" & total
            End If
        End If
        DoEvents
    Next
    
EARLY_OUT:
    Application.ScreenUpdating = True
    
    If pf.stopFlag Then
        pf.UpdateProgressBar 200, "�û���ֹ���Ѵ���" & done & "/" & total
        MsgBox "����ֹ���Ѵ���" & done & "/" & total & "�����б�� " & changed & " �Ρ�", vbExclamation
    Else
        pf.UpdateProgressBar 200, "��ɡ��ܼƣ�" & total & "����� " & changed & "������Ŀ����ʽ " & unchanged
        MsgBox "ͳһ��ɣ�" & vbCrLf & _
               "��ͼƬ����" & total & vbCrLf & _
               "����Ϊ��ͼƬ��ʽ����" & changed & vbCrLf & _
               "ԭ����Ϊ��ͼƬ��ʽ����" & unchanged, vbInformation
    End If
    
    Unload pf
End Sub

'=========================== ���ߺ��� ===========================

'��A���Ƿ�ΪͼƬ��InlineShape��
Private Function IsInlinePicture(ByVal ils As InlineShape) As Boolean
    On Error Resume Next
    Select Case ils.Type
        Case wdInlineShapePicture, wdInlineShapeLinkedPicture
            IsInlinePicture = True
        Case Else
            IsInlinePicture = False
    End Select
End Function

'��B���Ƿ�ΪͼƬ��Shape��������
Private Function IsFloatingPicture(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsFloatingPicture = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
End Function

'��C��ͳ������ Story �� Inline ͼƬ����
Private Function CountInlinePictures_MainStory(ByVal doc As Document) As Long
    Dim n As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        If IsInlinePicture(ils) Then
            If ils.Range.Paragraphs(1).Range.StoryType = wdMainTextStory Then n = n + 1
        End If
    Next
    CountInlinePictures_MainStory = n
End Function

'��D��ͳ������ Story �ĸ���ͼƬ����
Private Function CountFloatingPictures_MainStory(ByVal doc As Document) As Long
    Dim n As Long, s As Shape
    For Each s In doc.Shapes
        If IsFloatingPicture(s) Then
            If s.anchor.Paragraphs(1).Range.StoryType = wdMainTextStory Then n = n + 1
        End If
    Next
    CountFloatingPictures_MainStory = n
End Function

'��E�������Ƿ�����Ŀ����ʽ���ö���Ƚϸ��ȣ�
Private Function ParaHasStyle(ByVal p As Paragraph, ByVal sty As Style) As Boolean
    On Error Resume Next
    ParaHasStyle = (p.Range.Style Is sty)
End Function

'��F���ѡ���ǰ/����������� ProgressForm �����أ�0~200��
Private Function ProgressPixels(ByVal cur As Long, ByVal tot As Long) As Integer
    If tot <= 0 Then
        ProgressPixels = 0
    Else
        ProgressPixels = CInt(200# * cur / tot)
    End If
End Function


