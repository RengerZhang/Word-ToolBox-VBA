Attribute VB_Name = "���_4_�Զ�ɾ���ֹ����"
Option Explicit

'==========================================================
' �� �Զ�ɾ���ֹ���ţ��������У�
' ˵����
'   - Ŀ����ʽ��ɾ�����򣺾��� Mod�������� ��̬�ṩ
'   - ��ɾ�������ס��ֹ���ţ���Ӱ���Զ����
'   - ������һ�κ���ʽ������һ�Σ��������һ����ѭ��
'==========================================================
Sub ȥ���ֹ����_������������()

    Dim doc As Document
    Dim backupPath As String
    Dim targetStyles As Variant ' ��ʽ�����飨ֻȡ�ĵ��д��ڵģ�
    Dim patterns As Variant     ' ɾ���������飨���ڱ�Ÿ�ʽ��̬���ɣ�
    
    Dim styleName As Variant
    Dim rng As Range, contentRng As Range
    Dim originalText As String, newText As String
    Dim pat As Variant
    Dim matched As Boolean
    
    Dim lastStart As Long
    Dim nextPos As Long
    
    Set doc = ActiveDocument
    
    '��������ǰ�Զ����ݣ�ͬĿ¼��
    backupPath = ���ݵ�ǰ�ĵ�(doc)
    If Len(backupPath) > 0 Then Debug.Print "�ѱ��ݵ�: " & backupPath
    
    '��������������ȡ��ʽ & ���򣨲������
    targetStyles = ��ȡ��ʽ������(True)    ' ֻ�����ĵ��д��ڵ���ʽ
    patterns = ����ɾ����Ź���()         ' ���ڣ��ۣ��ı�Ÿ�ʽ�Զ��ƶ�
    
    '��������ÿһ��Ŀ����ʽ������ʽ Find����׼�Ҹ�Ч��
    For Each styleName In targetStyles
        
        Set rng = doc.content
        lastStart = -1
        
        With rng.Find
            .ClearFormatting
            .Style = doc.Styles(CStr(styleName))
            .text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            
            Do While .Execute
                '������ѭ�������������δ�ƽ�����ǿ������һ�Σ�
                If rng.Start = lastStart Then
                    nextPos = rng.Paragraphs(1).Range.End
                    If nextPos >= doc.content.End Then Exit Do
                    rng.SetRange Start:=nextPos, End:=doc.content.End
                    lastStart = -1
                    GoTo ContinueFindLoop
                End If
                lastStart = rng.Start
                
                '�������������ӷ�Χ��������β��ǣ�?��
                Set contentRng = rng.Duplicate
                If contentRng.Characters.Count > 1 Then
                    contentRng.MoveEnd wdCharacter, -1
                End If
                
                originalText = contentRng.text
                newText = originalText
                matched = False
                
                '�������� pass��������ɾ������������һ�飨�����ף�ÿ��ֻ��һ�Σ�
                For Each pat In patterns
                    If ��������(newText, CStr(pat)) Then
                        newText = �����滻(newText, CStr(pat), "")
                        matched = True
                    End If
                Next
                
                '�����������У�������ײ����ո񣨺�ȫ�ǣ�
                If matched Then
                    newText = �����滻(newText, "^[ ��]+", "")
                End If
                
                '���������б仯ʱд��
                If matched And newText <> originalText Then
                    contentRng.text = newText
                End If
                
                '�����ؼ�����ʽ������һ�Σ����׶ž�ĩ����ѭ��
                nextPos = rng.Paragraphs(1).Range.End
                If nextPos >= doc.content.End Then Exit Do
                rng.SetRange Start:=nextPos, End:=doc.content.End
                
ContinueFindLoop:
            Loop
        End With
    Next styleName
    
    MsgBox "�� ɾ���ֹ������ɣ���ʽ/������Զ���ȡ���ã���", vbInformation
End Sub

'�������������װ�����һ�£�
Private Function �����滻(ByVal s As String, ByVal pat As String, ByVal rep As String) As String
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    �����滻 = r.Replace(s, rep)
End Function

Private Function ��������(ByVal s As String, ByVal pat As String) As Boolean
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.pattern = pat
    �������� = r.TEST(s)
End Function

'�������ݣ�ͬ�����еĺ���һ�£��ɸ��ã�
Private Function ���ݵ�ǰ�ĵ�(ByVal doc As Document) As String
    On Error GoTo EH

    Dim baseName As String, ext As String, bak As String
    Dim folder As String, ts As String

    ts = Format(Now, "yyyymmdd_hhnnss")

    If Len(doc.name) > 0 Then
        baseName = doc.name
        If InStrRev(baseName, ".") > 0 Then
            ext = mid$(baseName, InStrRev(baseName, "."))
            baseName = Left$(baseName, InStrRev(doc.name, ".") - 1)
        Else
            ext = ".docx"
        End If
    Else
        baseName = "δ�����ĵ�"
        ext = ".docx"
    End If

    folder = IIf(doc.path = "", CurDir$, doc.path)
    If Right$(folder, 1) <> "\" Then folder = folder & "\"

    bak = folder & baseName & "_����_" & ts & ext
    doc.SaveCopyAs FileName:=bak

    ���ݵ�ǰ�ĵ� = bak
    Exit Function

EH:
    ���ݵ�ǰ�ĵ� = ""
End Function

