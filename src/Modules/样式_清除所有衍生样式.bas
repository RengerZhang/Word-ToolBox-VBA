Attribute VB_Name = "��ʽ_�������������ʽ"
Option Explicit

'========================================
' һ����ȫƪ������Ĭ����ʽ
'  - ������ʽ �� Normal��wdStyleNormal��
'  - �ַ���ʽ �� Default Paragraph Font��wdStyleDefaultParagraphFont��
'  - ɾ�����з����õĶ���/�ַ�/������ʽ
' ˵��������ֱ�Ӹ�ʽ���Ӵ�/б��/��ɫ�ȣ�����
'========================================
Public Sub ��ԭ��Ĭ����ʽ_�������Ĭ��()
    '��һ��׼�������ܿ���
    Dim doc As Document: Set doc = ActiveDocument
    Dim t0 As Single: t0 = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    '������ȡĬ����ʽ���������޹أ������ó������ȣ�
    Dim styParagraphDefault As Style
    Dim styCharacterDefault As Style
    Set styParagraphDefault = doc.Styles(wdStyleNormal)                 ' Normal / ����
    Set styCharacterDefault = doc.Styles(wdStyleDefaultParagraphFont)   ' Default Paragraph Font / Ĭ�϶�������
    
    '������A �������й��²� �� ������ʽͳһ��ԭΪ Normal
    Call ���²�_ȫ����Ϊ������ʽ(doc, styParagraphDefault)
    
    '���ģ�B ����������С���Ĭ�ϡ����ַ���ʽ����������/�Զ���/������ʽ���ַ��÷���
    '     ˼·����ÿһ�֡��ַ���������ʽ����ֻҪ���� Default Paragraph Font������ Find ȫ���滻Ϊ��
    Dim s As Style
    For Each s In doc.Styles
        If s.Type = wdStyleTypeCharacter Or s.Type = wdStyleTypeLinked Then
            If Not (s Is styCharacterDefault) And Not (s Is styParagraphDefault) Then
                Call ���²�_����ʽ�滻Ϊ(doc, s, styCharacterDefault)
            End If
        End If
    Next s
    
    '���壩��ѡ���������ֱ���ַ���ʽ��Ҳ��һ����������ɾ������ſ���һ�У�
    'doc.Content.ClearCharacterDirectFormatting
    
    '������C ����Ϊ���⡰��ʽ֮��ļ̳�/��һ����������ֹɾ�����Ȱѷ�������ʽ�������ĵ�Ĭ��
    Call ��ʽ_���������Ĭ��(doc, styParagraphDefault)
    
    '���ߣ�D ����ɾ��ȫ���������á��� ����/�ַ�/���� ��ʽ�����Ƿ��������޹أ�
    Call ɾ��������_�����ַ�������ʽ(doc)
    
    '���ˣ���β����ʾ
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    MsgBox "����ɣ�������Ĭ����ʽ" & vbCrLf & _
           "�� ������ʽ �� Normal" & vbCrLf & _
           "�� �ַ���ʽ �� Default Paragraph Font" & vbCrLf & _
           "�� ��������ʽ��ɾ��" & vbCrLf & _
           "�� ��ʱ���룩��" & Format$(Timer - t0, "0.0"), _
           vbInformation, "��ʽ��ԭ���"
End Sub

'========================================
'������һ�������й��²�Ķ�����ʽ����Ϊĳ������ʽ��ͨ���� Normal��
'========================================
Private Sub ���²�_ȫ����Ϊ������ʽ(doc As Document, ByVal paraStyle As Style)
    Dim rng As Range, r2 As Range
    For Each rng In doc.StoryRanges
        ' ��������Ӧ�ö�����ʽ���Զ�����Ч���ַ���ʽ������˱������
        rng.Style = paraStyle
        ' �����������²㣨�������ı�������
        Set r2 = rng
        Do While Not r2.NextStoryRange Is Nothing
            Set r2 = r2.NextStoryRange
            r2.Style = paraStyle
        Loop
    Next rng
End Sub

'========================================
'���������������й��²㣬�� oldS �� newS�������ַ�/������ʽ�滻��
'========================================
Private Sub ���²�_����ʽ�滻Ϊ(doc As Document, ByVal oldS As Style, ByVal newS As Style)
    Dim rng As Range, r2 As Range
    For Each rng In doc.StoryRanges
        Call ��Χ_����ʽ�滻(rng, oldS, newS)
        Set r2 = rng
        Do While Not r2.NextStoryRange Is Nothing
            Set r2 = r2.NextStoryRange
            Call ��Χ_����ʽ�滻(r2, oldS, newS)
        Loop
    Next rng
End Sub

Private Sub ��Χ_����ʽ�滻(ByVal rng As Range, ByVal oldS As Style, ByVal newS As Style)
    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .Style = oldS
        .replacement.Style = newS
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'========================================
'���������������С������á��� ����/�ַ�/���� ��ʽ�Ļ���ʽ/��һ����ʽ ָ�� Normal
' Ŀ�ģ���ֹ��ʽ֮�以Ϊ BaseStyle �� NextParagraphStyle ����ɾ��ʧ��
'========================================
Private Sub ��ʽ_���������Ĭ��(doc As Document, ByVal paraDefault As Style)
    Dim s As Style
    For Each s In doc.Styles
        On Error Resume Next
        If Not s.BuiltIn Then
            If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeLinked Then
                ' �������̳����롰��һ����ʽ���ĵ� Normal
                s.BaseStyle = paraDefault
                s.NextParagraphStyle = paraDefault
            End If
        End If
        On Error GoTo 0
    Next s
End Sub

'========================================
'�������ģ�ɾ�����С������á��� ����/�ַ�/���� ��ʽ
'========================================
Private Sub ɾ��������_�����ַ�������ʽ(doc As Document)
    Dim s As Style
    For Each s In doc.Styles
        If Not s.BuiltIn Then
            Select Case s.Type
                Case wdStyleTypeParagraph, wdStyleTypeCharacter, wdStyleTypeLinked
                    On Error Resume Next
                    s.Delete
                    Err.Clear
                    On Error GoTo 0
            End Select
        End If
    Next s
End Sub


