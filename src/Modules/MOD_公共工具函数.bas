Attribute VB_Name = "MOD_�������ߺ���"
Option Explicit

' ---------- ���ߺ����������ֺ��ı�ת��ֵ������������һ���� ----------
' ֧�������ֺţ��� ��š�С�� �ȣ���Ҳ����ֱ���������ְ�ֵ���� 12��10.5��
Public Function GetFontSizePt(sizeText As String) As Single
    Dim sizeMap As Object
    Set sizeMap = CreateObject("Scripting.Dictionary")
    With sizeMap
        .Add "����", 42
        .Add "С��", 36
        .Add "һ��", 26
        .Add "Сһ", 24
        .Add "����", 22
        .Add "С��", 18
        .Add "����", 16
        .Add "С��", 15
        .Add "�ĺ�", 14
        .Add "С��", 12
        .Add "���", 10.5
        .Add "С��", 9
        .Add "����", 7.5
        .Add "С��", 6.5
    End With

    If sizeMap.exists(sizeText) Then
        GetFontSizePt = sizeMap(sizeText)
    ElseIf IsNumeric(sizeText) Then
        GetFontSizePt = CSng(sizeText)
    Else
        GetFontSizePt = -1
    End If
End Function


'==========================================================
' �������ģ��ӡ��Զ���ţ��������������ȡ���������
' �Զ�������
'   1) Ŀ����ʽ�����飨����/��ʹ�ã�
'   2) ɾ���ֹ���ŵĹ��򼯣��������飬����ʹ�ã�
'   3) ����ƥ��Ĺ��򼯣�pattern��style ӳ�䣬����ʹ�ã�
'
' ������
'   - ��������� Public Function ��ȡ���м������() As Variant
'     ���ض�ά���飺(����, ��)���� = 1:��ʽ��, 2:��Ÿ�ʽ, 3:�����ʽ, 4:����λ��
'==========================================================

'������ ��ȡ����ʽ�������飨Ĭ��ֻ�����ĵ����Ѵ��ڵ���ʽ��
Public Function ��ȡ��ʽ������(Optional onlyExisting As Boolean = True) As Variant
    Dim cfg As Variant, i As Long, N As Long
    Dim buf() As String
    Dim sty As Style, name As String
    
    cfg = ��ȡ���м������()  ' ���ԡ�����������Ȩ��������
    ReDim buf(1 To UBound(cfg, 1))
    N = 0
    
    For i = 1 To UBound(cfg, 1)
        name = CStr(cfg(i, 1))
        If onlyExisting Then
            On Error Resume Next
            Set sty = ActiveDocument.Styles(name)
            On Error GoTo 0
            If Not sty Is Nothing Then
                N = N + 1: buf(N) = name
                Set sty = Nothing
            End If
        Else
            N = N + 1: buf(N) = name
        End If
    Next i
    
    If N = 0 Then
        ��ȡ��ʽ������ = Array() ' ��
    Else
        ReDim Preserve buf(1 To N)
        ��ȡ��ʽ������ = buf
    End If
End Function

'������ ɾ���ֹ���ţ����ݡ������ʽ+��Ÿ�ʽ����̬�������򣨽�ƥ����ף�
Public Function ����ɾ����Ź���() As Variant
    Dim cfg As Variant, i As Long, pat As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    cfg = ��ȡ���м������()
    For i = 1 To UBound(cfg, 1)
        pat = ����ɾ������(CLng(cfg(i, 3)), CStr(cfg(i, 2)))
        If Len(pat) > 0 Then
            If Not dict.exists(pat) Then dict.Add pat, True
        End If
    Next i
    
    ' ����������ţ�һ����������ʮһ�������ף�����ӿ�ѡ�����ո�
    ' ˵����ƥ�� 1~3 ���������֣��� ʮ/��/ǧ ��ϣ�������� ����.������: ��
    If Not dict.exists("^[ \t]*[һ�����������߰˾�ʮ��ǧ]{1,3}\s*(?:[��,��:������.\-���C]\s*)?") Then
        dict.Add "^[ \t]*[һ�����������߰˾�ʮ��ǧ]{1,3}\s*(?:[��,��:������.\-���C]\s*)?", True
    End If

    
    ' ����ͨ�ö��ף�������+�հף���ȫ�ǿո�
    If Not dict.exists("^\d+[ ��\t]+") Then dict.Add "^\d+[ ��\t]+", True
    
    If dict.Count = 0 Then
        ����ɾ����Ź��� = Array()
    Else
        ����ɾ����Ź��� = dict.Keys
    End If
End Function


'���� ɾ���ֹ���ţ������ݰ汾������ƥ�䣩
Private Function ����ɾ������(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = ռλ��(numFormat)
    Dim punct As String: punct = "[��,��:������.\-���C]"  ' ����ĺ�׺��㣨�ɰ���������

    ' 7 �����Ȧ�����֣���/?/�� �ȣ�
    If numStyle = wdListNumberStyleNumberInCircle Then
        ����ɾ������ = "^[ \t]*[" & ��������ż�() & "]\s*"
        Exit Function
    End If

    ' 6 ���%n���� (%n) ���� ȫ/������� + ����ո�
    If InStr(numFormat, "��%") > 0 Or InStr(numFormat, "(%") > 0 Then
        ����ɾ������ = "^[ \t]*[��(]\s*\d+\s*[)��]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' 5 ����%n���� %n) ���� ���� + �����ţ�ȫ/��ǣ�������ո�
    If Right$(Trim$(numFormat), 1) = "��" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            ����ɾ������ = "^[ \t]*\d+\s*[)��]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If

    ' �༶��1.1 �� 1�� 1��1 ���� ��ǰ����пո�ĩβ���ѡ
    If c >= 2 Then
        ����ɾ������ = "^[ \t]*\d+(?:\s*[\.����]\s*\d+){1,}\s*(?:[\.����])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' ������1 �� 1. ���� ���ų���1��/1)��������ǰհ��
    If c = 1 Then
        ����ɾ������ = "^[ \t]*\d+(?!\s*[)��])\s*(?:[\.����]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If

    ' ���ף������� + �հ�
    ����ɾ������ = "^[ \t]*\d+[ ��\t]+"
End Function
'====================���޸��棩�����ȼ����ɣ�����������4�Ρ�3�Ρ�2�Ρ�1��(����)��1��(������) ====================
Public Function ���ɱ���ƥ�����() As Variant
    Dim cfg As Variant, i As Long
    Dim sty As String, fmt As String, kind As Long
    Dim buckets(1 To 8) As Collection  ' ��һ��8�����ȼ�Ͱ
    Dim cat As Integer, pat As String
    Dim rowsCol As New Collection
    Dim rows() As Variant
    Dim p As Variant, j As Long, N As Long, k As Long
    Dim order As Variant
    
    '��������ʼ��8��Ͱ��1�� 2�� 3�� 4�Ķ� 5���� 6���� 7���δ��� 8���δ�����
    For i = 1 To 8
        Set buckets(i) = New Collection
    Next i
    
    '��������ȡ��Ų�����
    cfg = ��ȡ���м������()
    
    '���ģ���ÿ�����򶪵���Ӧ���ȼ�Ͱ��
    For i = 1 To UBound(cfg, 1)
        sty = CStr(cfg(i, 1))
        fmt = CStr(cfg(i, 2))
        kind = CLng(cfg(i, 3))
        
        pat = �������ƥ�����(kind, fmt)
        If Len(pat) > 0 Then
            cat = �������(kind, fmt)
            buckets(cat).Add Array(pat, sty)
        End If
    Next i
    
    '���壩���ȶ����ȼ�ƴ�ӣ����������ǰ��
    order = Array(1, 2, 3, 4, 5, 6, 7, 8)
    For Each p In order
        For j = 1 To buckets(p).Count
            rowsCol.Add buckets(p)(j)
        Next j
    Next p
    
    '������ת�ɶ�ά���鷵�ظ����÷�
    N = rowsCol.Count
    If N = 0 Then
        ���ɱ���ƥ����� = Array()
        Exit Function
    End If
    
    ReDim rows(1 To N, 1 To 2)
    For k = 1 To N
        rows(k, 1) = rowsCol(k)(0)  ' pattern
        rows(k, 2) = rowsCol(k)(1)  ' style
    Next k
    
    ���ɱ���ƥ����� = rows
End Function

'���������ף���һ�ֱ�Ÿ�ʽ���ൽ 8 �����ȼ�֮һ
Private Function �������(ByVal numStyle As Long, ByVal numFmt As String) As Integer
    Dim c As Long: c = ռλ��(numFmt)
    ' �� �Ȧ�ţ��� ���3
    If numStyle = wdListNumberStyleNumberInCircle Then
        ������� = 3: Exit Function
    End If
    ' �� ���n��/ (n)���� ���1
    If InStr(numFmt, "��%") > 0 Or InStr(numFmt, "(%") > 0 Then
        ������� = 1: Exit Function
    End If
    ' �� ����n��/ n)���� ���2
    If Right$(Trim$(numFmt), 1) = "��" Or Right$(Trim$(numFmt), 1) = ")" Then
        If InStr(numFmt, "%") > 0 Then ������� = 2: Exit Function
    End If
    ' �ܡ��� ���ʽ����
    Select Case c
        Case 4: ������� = 4: Exit Function
        Case 3: ������� = 5: Exit Function
        Case 2: ������� = 6: Exit Function
        Case 1
            If InStr(numFmt, ".") > 0 Or InStr(numFmt, "��") > 0 Or InStr(numFmt, "��") > 0 Then
                ������� = 7     ' ���δ��㣬�硰1.��
            Else
                ������� = 8     ' ���δ����֣��硰1��
            End If
        Case Else
            ������� = 8
    End Select
End Function

'====================���޸��棩�ϸ�ƥ�䣺����һ���������֡����� ====================
Private Function �������ƥ�����( _
    ByVal numStyle As Long, _
    ByVal numFormat As String _
) As String
    
    Dim c As Long: c = ռλ��(numFormat)
    Dim dot As String: dot = "[\.����]"   ' ����ĵ�ţ���/ȫ�ǣ�
    
    '��һ���ֻ��Ȧ��
    If numStyle = wdListNumberStyleNumberInCircle Then
        �������ƥ����� = "^[ \t]*[" & ��������ż�() & "]\s*"
        Exit Function
    End If
    
    '�������ֻ�ϣ�n��/ (n)
    If InStr(numFormat, "��%") > 0 Or InStr(numFormat, "(%") > 0 Then
        �������ƥ����� = "^[ \t]*[��(]\s*\d+\s*[)��]\s*"
        Exit Function
    End If
    
    '����������ֻ�� n��/ n)
    If Right$(Trim$(numFormat), 1) = "��" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            �������ƥ����� = "^[ \t]*\d+\s*[)��]\s*"
            Exit Function
        End If
    End If
    
    '���ģ����� 1~4���ϸ�ƥ�䡰ǡ�� N �Ρ�
    Select Case c
        Case 4
            �������ƥ����� = "^[ \t]*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 3
            �������ƥ����� = "^[ \t]*\d+\s*" & dot & "\s*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 2
            �������ƥ����� = "^[ \t]*\d+\s*" & dot & "\s*\d+(?!\s*" & dot & "\s*\d)"
        Case 1
            ' ���һ����ʽ�� ��%1  ���������㣬���������� ���ų� ��1��/1)�� �� ��1.1�� ���ָ���
            '���ؼ�����������ǰհ���÷��飬һ���ų�����ǰ׺�������� �� ����+���֡�
            If InStr(numFormat, ".") > 0 Or InStr(numFormat, "��") > 0 Or InStr(numFormat, "��") > 0 Then
                �������ƥ����� = "^[ \t]*\d+\s*" & dot & "(?!\s*\d)"
            Else
                �������ƥ����� = "^[ \t]*\d+(?!\s*(?:[)��]|" & dot & "\s*\d))"
            End If
        Case Else
            �������ƥ����� = ""
    End Select
End Function



'���� ����������Ȧ�����ּ��ϣ�������֡�����������
Public Function ��������ż�() As String
    Dim s As String, code As Long
    For code = &H2460 To &H2473: s = s & ChrW(code): Next code   ' ��..?
    For code = &H2474 To &H2487: s = s & ChrW(code): Next code   ' ��..��
    For code = &H2776 To &H277F: s = s & ChrW(code): Next code   ' ?..?
    For code = &H24EB To &H24F4: s = s & ChrW(code): Next code   ' ?..?
    For code = &H24F5 To &H24FE: s = s & ChrW(code): Next code   ' ?..?
    ��������ż� = s
End Function


' ͳ�� "%n" ռλ������
Private Function ռλ��(ByVal fmt As String) As Long
    Dim i As Long, c As Long
    For i = 1 To Len(fmt)
        If mid$(fmt, i, 1) = "%" Then c = c + 1
    Next
    ռλ�� = c
End Function

'������ ��ȡ����ʽ�������飨���ָ���ĵ���onlyExisting=True ʱֻ���ظ��ĵ�����ڵ���ʽ��
Public Function ��ȡ��ʽ������_����ĵ�(ByVal src As Document, Optional onlyExisting As Boolean = True) As Variant
    Dim cfg As Variant, i As Long, N As Long
    Dim buf() As String
    Dim sty As Style, name As String

    cfg = ��ȡ���м������()  ' ���Բ��裨����
    ReDim buf(1 To UBound(cfg, 1))
    N = 0

    For i = 1 To UBound(cfg, 1)
        name = CStr(cfg(i, 1))
        If onlyExisting Then
            On Error Resume Next
            Set sty = src.Styles(name)
            On Error GoTo 0
            If Not sty Is Nothing Then
                N = N + 1: buf(N) = name
                Set sty = Nothing
            End If
        Else
            N = N + 1: buf(N) = name
        End If
    Next i

    If N = 0 Then
        ��ȡ��ʽ������_����ĵ� = Array()
    Else
        ReDim Preserve buf(1 To N)
        ��ȡ��ʽ������_����ĵ� = buf
    End If
End Function

