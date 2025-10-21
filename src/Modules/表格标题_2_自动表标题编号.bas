Attribute VB_Name = "������_2_�Զ��������"
Option Explicit

' ==========================================================
' �������Զ���ţ�������������ɾ���ֹ���ţ�
' ����
'   - ��� = ���� + �������ţ������ļ���������������һ���� + ��-�� + ͬһ�����������������
'   - ��ʾ�ţ����ļ����ļ� a.b.c.d�����������м�����a / a.b / a.b.c��
'   - ��ŷ�������̶��������� a.b.c����������ʱ�����м�����
'   - �����������еı���ǰ���ж���հ׶Σ������ҵ�һ���ǿն���Ϊ����Σ�
'   - �������Ѵ��ڡ��� + ��š�ǰ׺���������������ǡ���ɾ����
'   - ���ȴ��壺ProgressForm�����Ѵ��� UpdateProgressBar/stopFlag �ȳ�Ա��
' ==========================================================

Public Sub �������Զ����_ʹ�ý��ȴ���1()
    Dim doc As Document: Set doc = ActiveDocument
    Dim totalTables As Long, done As Long, passMsg As String
    Dim ��ż��� As Object: Set ��ż��� = CreateObject("Scripting.Dictionary")
    Dim tbl As Table, tblIdx As Long

    ' ͳ�������еı����������ڽ�����������
    totalTables = ͳ�����ı�����()

    ' �򿪽��ȴ���
    With progressForm
        .caption = "����Զ����"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = "��ʼ���� " & totalTables & " ����" & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "����Զ����"
    On Error GoTo 0

    tblIdx = 0
    For Each tbl In doc.Tables
        If tbl.Range.StoryType <> wdMainTextStory Then GoTo NextTable
        tblIdx = tblIdx + 1
        If progressForm.stopFlag Then Exit For

        ' ��Ŵ���������
        �������� tbl, ��ż���, tblIdx, totalTables

        ' ����
        done = tblIdx
        progressForm.UpdateProgressBar ��ǰ��������(done, IIf(totalTables = 0, 1, totalTables)), _
            "���ȣ�" & done & "/" & totalTables
        DoEvents

NextTable:
    Next tbl

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

    If Not progressForm.stopFlag Then
        progressForm.UpdateProgressBar 200, "��ɡ�"
        MsgBox "��������д���ţ���ɾ���ֹ���ţ���" & vbCrLf & _
               "��ʾ��Ctrl+G �򿪡��������ڡ��鿴��ϸ��־��", vbInformation
    Else
        MsgBox "���ֶ���ֹ��", vbExclamation
    End If

    On Error Resume Next
    Unload progressForm
    On Error GoTo 0
End Sub


' ----------------------------------------------------------
' ����������λ����Ρ���λ��������������ʾ��/����Key��д����
' ----------------------------------------------------------
Private Sub ��������(ByVal tbl As Table, _
                      ByRef ��ż��� As Object, _
                      ByVal ���� As Long, _
                      ByVal totalTables As Long)

    Const ������ʽ�� As String = "������"

    Dim doc As Document: Set doc = ActiveDocument
    Dim tblRng As Range, prevPara As Paragraph, paraText As String
    Dim h As Range, ����Ԥ�� As String
    Dim ���� As Long, ԭList As String, ������ As String
    Dim ������ As Variant, ��ʾ�� As String, ����Key As String
    Dim segDump As String
    Dim r As Range, ���� As String, ������ As String

    Set tblRng = tbl.Range.Duplicate

    ' ������ǰ��һ���ǿն���Ϊ����Σ��������հ׶Σ�
    Set prevPara = ����ȡ��һ���ǿն�(tblRng)
    If prevPara Is Nothing Then
        progressForm.UpdateProgressBar ��ǰ��������(����, IIf(totalTables = 0, 1, totalTables)), _
            "��#" & ���� & "��δ�ҵ�����Σ�������"
        Exit Sub
    End If

    ' �������á������⡱��ʽ�����Ѵ����򲻱䣩
    On Error Resume Next
    prevPara.Style = doc.Styles(������ʽ��)
    On Error GoTo 0

    ' �������С��� + ��š�ǰ׺�������������ǣ�
    paraText = ������׿ɼ��ı�(prevPara.Range.text)
    If ��������(paraText, "^\s*��\s*\d+(?:[\.����]\s*\d+){0,6}\s*[-���C��]\s*\d+") Then
        progressForm.UpdateProgressBar ��ǰ��������(����, IIf(totalTables = 0, 1, totalTables)), _
            "��#" & ���� & "����⵽���б�ţ��������� " & Left$(paraText, 40)
        Exit Sub
    End If

    ' ������λ����½ڱ��⣨�����ļ���������������һ����
    Set h = ��λ�������_GoTo(prevPara.Range)
    If Not h Is Nothing Then
        ���� = h.Paragraphs(1).outlineLevel
        On Error Resume Next
        ԭList = h.Paragraphs(1).Range.ListFormat.ListString
        On Error GoTo 0

        ������ = ��ȡ��׼��Ŵ�(h.Paragraphs(1))       ' �� "3.1.4.1"
        ������ = ��ȡ��Ŷ�����(������)                ' Array("3","1","4","1") �� Empty
        ��ʾ�� = ������ʾ��_����ļ�(������)           ' �� 4 ���� 4 �Σ����������ж�
        ����Key = �������Key_��������(������)         ' ʼ�հ�����������
        ����Ԥ�� = ������׿ɼ��ı�(h.Paragraphs(1).Range.text)
    Else
        ���� = 0: ԭList = "": ������ = "": ��ʾ�� = "": ����Key = ""
        ����Ԥ�� = "(δ�ҵ�����)"
    End If

    ' �����������
    If IsArray(������) Then
        On Error Resume Next
        segDump = Join(������, ",")
        If Err.Number <> 0 Then segDump = "(��)": Err.Clear
        On Error GoTo 0
    Else
        segDump = "(��)"
    End If

    Debug.Print "��#" & ���� & "������=" & ���� & _
                " | ListString=[" & ԭList & "]" & _
                " | ������=[" & ������ & "]" & _
                " | ������=(" & segDump & ")" & _
                " | ��ʾ��=[" & ��ʾ�� & "]" & _
                " | ����Key=[" & ����Key & "]" & _
                " | ����� "; Left$(����Ԥ��, 40)

    ' ����������ͬһ���������������ۼӣ�
    If Len(����Key) = 0 Then ����Key = "0"
    If Not ��ż���.exists(����Key) Then ��ż���.Add ����Key, 0
    ��ż���(����Key) = ��ż���(����Key) + 1

    ' ����д���ţ�����ʽ�������ǰ׺����д��ǰ׺��
    Set r = prevPara.Range.Duplicate
    If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1

    ���� = r.text
    ������ = �����滻_һ��(����, "^\s*��\s*\d+(?:[\.����]\s*\d+){0,6}\s*[-���C��]\s*\d+\s*", "")
    ������ = LTrim$(������)

    r.text = "��" & ��ʾ�� & "-" & CStr(��ż���(����Key)) & "  " & ������

    progressForm.UpdateProgressBar ��ǰ��������(����, IIf(totalTables = 0, 1, totalTables)), _
        "��#" & ���� & "��д�� �� ��" & ��ʾ�� & "-" & ��ż���(����Key)
End Sub


' ========================= �����빤�� =========================

' ͳ�������б������
Private Function ͳ�����ı�����() As Long
    Dim t As Table
    Dim n As Long
    For Each t In ActiveDocument.Tables
        If t.Range.StoryType = wdMainTextStory Then n = n + 1
    Next
    ͳ�����ı����� = n
End Function

' ����ȡ��һ���ǿնΣ��������հ׶Σ�
Private Function ����ȡ��һ���ǿն�(ByVal tblRng As Range) As Paragraph
    Dim p As Paragraph, s As String
    If tblRng.Paragraphs.Count = 0 Then Exit Function
    Set p = tblRng.Paragraphs(1).Previous
    Do While Not p Is Nothing
        s = ������׿ɼ��ı�(p.Range.text)
        If Len(s) > 0 Then Set ����ȡ��һ���ǿն� = p: Exit Function
        Set p = p.Previous
    Loop
End Function

' ��ê�����ϣ������ͽ�+��λ������λ�������
' �����ļ�����������������������ļ������򷵻ظ�������������/һ��ֱ�ӷ���
Private Function ��λ�������_GoTo(ByVal anchor As Range) As Range
    Dim base As Range, cur As Range, hop As Range
    Dim cand4 As Range, lvl As Long, guard As Long

    Set base = anchor.Duplicate
    base.SetRange Start:=base.Start, End:=base.Start

    Set cur = base.Duplicate
    Do
        On Error Resume Next
        Set hop = cur.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        On Error GoTo 0
        If hop Is Nothing Then Exit Do
        If hop.Start >= cur.Start Then Exit Do  ' ����ѭ��

        Set cur = hop
        lvl = cur.Paragraphs(1).outlineLevel

        Select Case lvl
            Case wdOutlineLevel4
                If cand4 Is Nothing Then Set cand4 = cur.Paragraphs(1).Range
            Case wdOutlineLevel3
                If Not cand4 Is Nothing Then
                    Set ��λ�������_GoTo = cand4
                Else
                    Set ��λ�������_GoTo = cur.Paragraphs(1).Range
                End If
                Exit Function
            Case wdOutlineLevel2, wdOutlineLevel1
                Set ��λ�������_GoTo = cur.Paragraphs(1).Range
                Exit Function
            Case Else
                ' ���������������
        End Select

        guard = guard + 1
        If guard > 20000 Then Exit Do
    Loop

    If Not cand4 Is Nothing Then Set ��λ�������_GoTo = cand4
End Function

' �ѱ����תΪ��׼��Ŵ������� ListString��ʧ����Ӷ����ı�������
Private Function ��ȡ��׼��Ŵ�(ByVal p As Paragraph) As String
    Dim s As String, t As String
    On Error Resume Next
    s = p.Range.ListFormat.ListString
    On Error GoTo 0
    s = �淶����Ŵ�(s)
    If Len(s) > 0 Then
        ��ȡ��׼��Ŵ� = s
        Exit Function
    End If
    t = �������ױ��(p.Range.text)
    ��ȡ��׼��Ŵ� = t
End Function

' ��ȡ��Ŷ����飺ֻ����������㣬ѹ����㣬Split
Private Function ��ȡ��Ŷ�����(ByVal numStr As String) As Variant
    Dim s As String
    s = Replace$(Replace$(numStr, "��", "."), "��", ".")
    s = �����滻_ȫ��(s, "[^\d\.]", "")
    s = �����滻_ȫ��(s, "^\.+|\.+$", "")
    s = �����滻_ȫ��(s, "\.+", ".")
    If Len(s) = 0 Then
        ��ȡ��Ŷ����� = Empty
    Else
        ��ȡ��Ŷ����� = Split(s, ".")
    End If
End Function

' ��ʾ�ţ��� 4 ���� 4 �Σ��������ж�����1/2/3��
Private Function ������ʾ��_����ļ�(ByVal segs As Variant) As String
    Dim n As Long
    If IsEmpty(segs) Then Exit Function
    n = UBound(segs) - LBound(segs) + 1
    Select Case n
        Case Is >= 4: ������ʾ��_����ļ� = segs(0) & "." & segs(1) & "." & segs(2) & "." & segs(3)
        Case 3:       ������ʾ��_����ļ� = segs(0) & "." & segs(1) & "." & segs(2)
        Case 2:       ������ʾ��_����ļ� = segs(0) & "." & segs(1)
        Case Else:    ������ʾ��_����ļ� = segs(0)
    End Select
End Function

' ����Key���̶��õ�����������������ʱ�����ж���
Private Function �������Key_��������(ByVal segs As Variant) As String
    Dim n As Long
    If IsEmpty(segs) Then Exit Function
    n = UBound(segs) - LBound(segs) + 1
    Select Case n
        Case Is >= 3: �������Key_�������� = segs(0) & "." & segs(1) & "." & segs(2)
        Case 2:       �������Key_�������� = segs(0) & "." & segs(1)
        Case Else:    �������Key_�������� = segs(0)
    End Select
End Function

' �淶����ţ�ȥ�հ�/ȫ�ǵ����ǣ�������������㣻ѹ����㣻ȥ��β�㣻����У��
Private Function �淶����Ŵ�(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function
    s = Replace$(s, vbCr, "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = Replace$(s, "��", ".")
    s = Replace$(s, "��", ".")
    s = �����滻_ȫ��(s, "\s+", "")
    s = �����滻_ȫ��(s, "[^\d\.]", "")
    s = �����滻_ȫ��(s, "\.+", ".")
    s = �����滻_ȫ��(s, "^\.|\.?$", "")
    If ��������(s, "^\d+(?:\.\d+){0,7}$") Then �淶����Ŵ� = s
End Function

' �Ӷ����ı�������ţ������ǰ��ո�/ȫ�ǵ㣩
Private Function �������ױ��(ByVal s As String) As String
    Dim m As Object
    s = Replace$(Replace$(s, "��", "."), "��", ".")
    s = Replace$(s, vbCr, "")
    Set m = ����ƥ��(s, "^\s*\d+(?:\s*\.\s*\d+){0,7}")
    If Not m Is Nothing Then �������ױ�� = �淶����Ŵ�(m.Value)
End Function

' ȥ��β/��Ԫ�������/ȫ�ǿո����ǣ��� Trim
Private Function ������׿ɼ��ı�(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    ������׿ɼ��ı� = Trim$(s)
End Function

' �������򹤾�
Private Function ��������(ByVal s As String, ByVal pat As String) As Boolean
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = False: rx.Global = False: rx.pattern = pat
    �������� = rx.TEST(s)
End Function

Private Function ����ƥ��(ByVal s As String, ByVal pat As String) As Object
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    Dim mc As Object
    rx.IgnoreCase = False: rx.Global = False: rx.pattern = pat
    Set mc = rx.Execute(s)
    If mc.Count > 0 Then Set ����ƥ�� = mc(0) Else Set ����ƥ�� = Nothing
End Function

' �����滻������ɾ�����С���+��š�ǰ׺��ֻ���״���������ɾ���ģ�
Private Function �����滻_һ��(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    �����滻_һ�� = rx.Replace(s, rep)
End Function

' ȫ���滻
Private Function �����滻_ȫ��(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = False: rx.Global = True: rx.pattern = pat
    �����滻_ȫ�� = rx.Replace(s, rep)
End Function

' �������أ���Ĵ������� 200px��
Private Function ��ǰ��������(ByVal done As Long, ByVal total As Long) As Long
    If total <= 0 Then
        ��ǰ�������� = 0
    Else
        ��ǰ�������� = CLng(200# * done / total)
        If ��ǰ�������� < 0 Then ��ǰ�������� = 0
        If ��ǰ�������� > 200 Then ��ǰ�������� = 200
    End If
End Function


