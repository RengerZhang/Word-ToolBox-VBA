Attribute VB_Name = "���_2_�Զ�ƥ���������"
Option Explicit

'==========================================================
' �� ������ʽ�Զ�����ƥ�䣨�������У�
' ˵����
'   - ͳһ�� Mod�������� ��ȡ pattern��style �Ĺ���
'   - �ѱ�Ŷ������ȣ��� ListLevelNumber ӳ�䵽��ʽ
'   - δ��Ŷ��䣺�ö�̬����ƥ����ױ����̬���Զ����ö�Ӧ��ʽ
'   - �����ı�ţ�ֻ�ı������ʽ
'==========================================================
Sub ƥ����Ⲣ������ʽ_������������()

    Dim doc As Document
    Dim rules As Variant ' [[pattern, style], ...]
    Dim cfg As Variant   ' ���ԣ��ۣ�����ȡ���м������()����[��, ��1..4]
    Dim level2Style() As String
    Dim Para As Paragraph
    Dim t As String
    Dim lvl As Long
    Dim i As Long
    Dim tocZones As Collection
    Set tocZones = ����TOC����(doc)
    
    Set doc = ActiveDocument
    
    '�������򣺴��������ġ�����Ÿ�ʽ����̬���ɣ�������д��
    rules = ���ɱ���ƥ�����()
    
    '������������ż��� �� ��ʽ����ӳ�䣨���Զ���Ŷ���ֱ������ʽ��
    cfg = ��ȡ���м������()
    ReDim level2Style(1 To UBound(cfg, 1))
    For i = 1 To UBound(cfg, 1)
        level2Style(i) = CStr(cfg(i, 1)) ' ��1�У���ʽ��
    Next i
    
     '������δ���
    For Each Para In doc.Paragraphs
    ' 0) �����Ĺ��£��ų�ҳüҳ�š��ı����
    If Para.Range.StoryType <> wdMainTextStory Then GoTo NextPara

    ' 0.1) �ų�������еĶ��䡱��˫���գ�Information + Tables.Count��
    Dim inTable As Boolean
    On Error Resume Next
    inTable = Para.Range.Information(wdWithInTable)   ' True ��ʾ�ڱ����
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If Not inTable Then inTable = (Para.Range.Tables.Count > 0)
    If inTable Then GoTo NextPara
    
    ' 0.2) �ų���Ŀ¼�еĶ��䡱
    '     �� ��ʽ���� "TOC" �� "Ŀ¼" ��ͷ
    '     �� ���߶���� Range ������һ TOC �ֶεĽ��������
    Dim sty As String
    On Error Resume Next
    sty = Para.Range.Style.nameLocal
    On Error GoTo 0
    If Len(sty) > 0 Then
        If (UCase$(Left$(sty, 3)) = "TOC" Or Left$(sty, 2) = "Ŀ¼") Then GoTo NextPara
    End If
    If ��TOC������(Para.Range, tocZones) Then GoTo NextPara


    ' 1) ��ȡ������ɼ��ı�
    t = ��������ı�(Para.Range.text)
    If Len(t) = 0 Then GoTo NextPara

    ' 2) �Զ�������ȣ�ֱ�Ӱ�����ӳ��
    On Error Resume Next
    If Para.Range.ListFormat.ListType <> wdListNoNumbering Then
        lvl = Para.Range.ListFormat.ListLevelNumber
    Else
        lvl = 0
    End If
    On Error GoTo 0

    If lvl >= LBound(level2Style) And lvl <= UBound(level2Style) Then
        If ��ʽ����(doc, level2Style(lvl)) Then
            Para.Style = doc.Styles(level2Style(lvl))
            GoTo NextPara
        End If
    End If

    ' 3) ���Զ���ţ��ö�̬����ƥ����ױ��ģʽ �� ����ʽ
    If IsArray(rules) Then
        For i = LBound(rules, 1) To UBound(rules, 1)
            If ��������(t, CStr(rules(i, 1))) Then
                If ��ʽ����(doc, CStr(rules(i, 2))) Then
                    Para.Style = doc.Styles(CStr(rules(i, 2)))
                    Exit For
                End If
            End If
        Next i
    End If

NextPara:
Next Para
    
    MsgBox "�ڱ���ƥ����ɣ���", vbInformation
End Sub

'�������ߣ���ʽ�Ƿ����
Private Function ��ʽ����(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim s As Style
    On Error Resume Next
    Set s = doc.Styles(styleName)
    ��ʽ���� = Not (s Is Nothing)
    Set s = Nothing
    On Error GoTo 0
End Function

'�������ߣ�������ԣ����ж���
Private Function ��������(ByVal s As String, ByVal pat As String) As Boolean
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    �������� = r.TEST(s)
End Function

'�������ߣ��������ɼ��ı���ȥ��β���/��Ԫ�������/ȫ�ǿո����ǡ�Trim��
Private Function ��������ı�(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")              ' ��Ԫ�������
    s = Replace$(s, ChrW(&H3000), " ")       ' ȫ�ǿո�����
    ��������ı� = Trim$(s)
End Function
'�������� TOC �ֶν�����򼯺ϣ���� text �������䣬��������룩
Private Function ����TOC����(ByVal doc As Document) As Collection
    Dim zones As New Collection
    Dim f As Field, codeTxt As String
    On Error Resume Next
    For Each f In doc.Fields
        ' ���ֶ����ͻ�����ı��ж���������ؼ���ʼ���� "TOC"��
        codeTxt = ""
        codeTxt = f.code.text
        If (f.Type = wdFieldTOC) Or (InStr(1, UCase$(codeTxt), "TOC", vbTextCompare) > 0) Then
            zones.Add f.Result.Duplicate
        End If
    Next f
    Set ����TOC���� = zones
End Function

'�����ж�һ�� Range �Ƿ���ȫ��������һ�� TOC ���������
Private Function ��TOC������(ByVal r As Range, ByVal zones As Collection) As Boolean
    Dim z As Range
    If zones Is Nothing Then Exit Function
    On Error Resume Next
    For Each z In zones
        If (r.Start >= z.Start) And (r.End <= z.End) Then
            ��TOC������ = True
            Exit Function
        End If
    Next z
End Function


