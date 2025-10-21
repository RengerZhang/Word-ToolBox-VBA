Attribute VB_Name = "���_6_�����ȵ�ȥ�����"
Option Explicit

'==========================================================
' �� ɾ���ֹ���ţ��������� + ѭ��ȷ�ϣ�
' ˵����
'   - Ŀ����ʽ & ɾ���������ԡ��������ġ�
'   - ��ɾ�������ֹ���ţ���Ӱ��༶�Զ����
'   - ����ʱ������������Story������ڶ��䡢Ŀ¼(TOC)����
'   - ÿһ�ֽ���������������>0����ѯ���Ƿ������һ��
'==========================================================
Public Sub ȥ���ֹ����_ʹ�ý��ȴ���()
    Dim doc As Document: Set doc = ActiveDocument
    Dim scope As Range              ' �� ���δ���Χ��ѡ�� �� ȫ��
    Dim scopeInfo As String
    Dim backupPath As String
    Dim targetStyles As Variant
    Dim patterns As Variant
    Dim tocZones As Collection
    Dim cand As Long
    Dim passNo As Long
    Dim touched As Long, skipped As Long, total As Long, allTouched As Long
    Dim ans As VbMsgBoxResult

    '����1) ����Χ������ѡ����������Story����������ѡ�ж��䣻����ȫ��
    If Selection.Type <> wdSelectionIP And Selection.Range.StoryType = wdMainTextStory Then
        Set scope = Selection.Range.Duplicate
        scopeInfo = "��Χ��ѡ�ж���"
    Else
        Set scope = doc.content.Duplicate
        scopeInfo = "��Χ��ȫ��"
    End If

    '����2) ����ǰ���ݣ�ͬĿ¼��
    backupPath = ���ݵ�ǰ�ĵ�(doc)
    If Len(backupPath) > 0 Then Debug.Print "�ѱ��ݵ�: " & backupPath

    '����3) �ӡ��������ġ���ȡ����ʽ���� & ɾ������
    targetStyles = ��ȡ��ʽ������(True)          ' ֻ�����ĵ����Ѵ��ڵ�Ŀ����ʽ
    patterns = ����ɾ����Ź���()               ' ��̬����ɾ��ģʽ

    '����4) �ռ� TOC ������������Ŀ¼���ݣ�
    Set tocZones = ����TOC����(doc)

    '����5) ͳ�ƺ�ѡ�Σ��� scope �ڣ�
    cand = ͳ�ƺ�ѡ����_ɾ�����(scope, targetStyles, tocZones)

    '����6) �򿪽��ȴ���
    With progressForm
        .caption = "ɾ���ֹ����"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = scopeInfo & "����ѡ���䣺" & cand & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "ɾ���ֹ���ţ�ѭ����"
    On Error GoTo 0

    Do
        passNo = passNo + 1
        progressForm.UpdateProgressBar 0, "���� �� " & passNo & " �ֿ�ʼ ����"

        ' ÿ�ֿ�ͷ����ͳ�ƣ���һ�ֿ��ܸĶ�����ʽ/�ı���
        cand = ͳ�ƺ�ѡ����_ɾ�����(scope, targetStyles, tocZones)
        If cand = 0 Then
            progressForm.UpdateProgressBar 200, "û�к�ѡ���䣬ֱ�ӽ�����"
            Exit Do
        End If

        total = 0: touched = 0: skipped = 0
        ִ��һ��ɾ�� scope, targetStyles, patterns, tocZones, cand, total, touched, skipped

        progressForm.UpdateProgressBar 200, _
            "�� " & passNo & " ��С�᣺��ѡ=" & total & "�������=" & touched & "��δ���=" & skipped

        allTouched = allTouched + touched
        If progressForm.stopFlag Then
            MsgBox "���ֶ���ֹ���ۼ������" & allTouched & " ����", vbExclamation
            Exit Do
        End If

        If touched = 0 Then
            MsgBox "ɾ���ֹ������ɣ������޿�������" & vbCrLf & _
                   "�ۼ������" & allTouched & " ����", vbInformation
            Exit Do
        Else
            ans = MsgBox("��������� " & touched & " �����ǰ׺��" & vbCrLf & _
                         "�������в������Ƿ������һ�֣�", _
                         vbYesNo + vbQuestion, "���������")
            If ans = vbNo Then Exit Do
            progressForm.FrameProgress.width = 0
            progressForm.LabelPercentage.caption = "0%"
        End If

        If passNo >= 5 Then
            ans = MsgBox("������ 5 �֣��Ƿ��Լ�����", vbYesNo + vbExclamation)
            If ans = vbNo Then Exit Do
        End If
    Loop

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

    If Not progressForm.stopFlag Then
        progressForm.UpdateProgressBar 200, "��ɡ��ۼ������" & allTouched
        MsgBox "ɾ���ֹ���ţ�����������ۼ���� " & allTouched & " ����", vbInformation
    End If
End Sub



' ----------------------------------------------------------
' ����ִ�У�������Ŀ����ʽ��ɾ�����ԣ������ף�
'  - �����������ġ�����ڡ�TOC ����
'  - ÿ����һ�Σ�����������ɾ�����򣨵����滻����������׿ո�
'  - Ϊ��ֹ��ĩ����ѭ������ÿ����ʽ�� rng ������һ��
' candTotal ���ڽ���������� total/touched/skipped
' ----------------------------------------------------------
' ����ִ�У���α����������� Find������ĩ��Խ��
Private Sub ִ��һ��ɾ��(ByVal scope As Range, _
                       ByVal targetStyles As Variant, _
                       ByVal patterns As Variant, _
                       ByVal tocZones As Collection, _
                       ByVal candTotal As Long, _
                       ByRef total As Long, _
                       ByRef touched As Long, _
                       ByRef skipped As Long)

    Dim p As Paragraph
    Dim contentRng As Range
    Dim originalText As String, newText As String
    Dim pat As Variant, sty As String
    Dim processed As Long, examples As Long

    For Each p In scope.Paragraphs
        ' ���ˣ����� / �Ǳ�� / ��Ŀ¼
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextP
        If �ڱ����(p.Range) Then GoTo NextP
        If ��TOC������(p.Range, tocZones) Then GoTo NextP

        On Error Resume Next
        sty = p.Range.Style.nameLocal
        On Error GoTo 0
        If Not ��ʽ���б���(sty, targetStyles) Then GoTo NextP

        total = total + 1

        ' ȡ�ɱ༭���ݣ�������β��ǣ�
        Set contentRng = p.Range.Duplicate
        If contentRng.Characters.Count > 1 Then contentRng.MoveEnd wdCharacter, -1

        originalText = contentRng.text
        newText = originalText

        ' �����������С�ɾ�����ױ�š�������ÿ��ֻ���״���
        For Each pat In patterns
            newText = �����滻_һ��(newText, CStr(pat), "")
        Next
        ' ����ײ����ո񣨺�ȫ�ǣ�
        newText = �����滻_һ��(newText, "^[ ��]+", "")

        If newText <> originalText Then
            contentRng.text = newText
            touched = touched + 1
            If examples < 6 Then
                progressForm.UpdateProgressBar ��ǰ��������(processed, IIf(candTotal = 0, 1, candTotal)), _
                    "���ǰ��" & Left$(originalText, 80) & vbCrLf & " �ĺ�" & Left$(newText, 80)
                examples = examples + 1
            End If
        Else
            skipped = skipped + 1
        End If

        processed = processed + 1
        progressForm.UpdateProgressBar ��ǰ��������(processed, IIf(candTotal = 0, 1, candTotal)), _
            "���ȣ�" & processed & "/" & candTotal
NextP:
        DoEvents
    Next p
End Sub


' ��ѡ�Σ�����Story���Ǳ�񡢷�TOC����ʽ �� Ŀ����ʽ��
Private Function ͳ�ƺ�ѡ����_ɾ�����(ByVal scope As Range, _
                                    ByVal targetStyles As Variant, _
                                    ByVal tocZones As Collection) As Long
    Dim p As Paragraph, n As Long, sty As String

    For Each p In scope.Paragraphs
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextP
        If �ڱ����(p.Range) Then GoTo NextP
        If ��TOC������(p.Range, tocZones) Then GoTo NextP
        On Error Resume Next
        sty = p.Range.Style.nameLocal
        On Error GoTo 0
        If ��ʽ���б���(sty, targetStyles) Then n = n + 1
NextP:
    Next
    ͳ�ƺ�ѡ����_ɾ����� = n
End Function

Private Function ��ʽ���б���(ByVal sty As String, ByVal arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If StrComp(sty, CStr(v), vbTextCompare) = 0 Then ��ʽ���б��� = True: Exit Function
    Next
End Function

' �Ƿ��ڱ���˫���գ�
Private Function �ڱ����(ByVal r As Range) As Boolean
    On Error Resume Next
    �ڱ���� = r.Information(wdWithInTable)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If Not �ڱ���� Then �ڱ���� = (r.Tables.Count > 0)
End Function

'�������� TOC �ֶν�����򼯺�
Private Function ����TOC����(ByVal doc As Document) As Collection
    Dim zones As New Collection
    Dim f As Field, codeTxt As String
    On Error Resume Next
    For Each f In doc.Fields
        codeTxt = ""
        codeTxt = f.code.text
        If (f.Type = wdFieldTOC) Or (InStr(1, UCase$(codeTxt), "TOC", vbTextCompare) > 0) Then
            zones.Add f.Result.Duplicate
        End If
    Next f
    Set ����TOC���� = zones
End Function

'�����ж� Range �Ƿ���ȫ������һ TOC ���������
Private Function ��TOC������(ByVal r As Range, ByVal zones As Collection) As Boolean
    Dim z As Range
    If zones Is Nothing Then Exit Function
    On Error Resume Next
    For Each z In zones
        If (r.Start >= z.Start) And (r.End <= z.End) Then ��TOC������ = True: Exit Function
    Next z
End Function

' ���򣺵����滻�����״���
Private Function �����滻_һ��(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    �����滻_һ�� = rx.Replace(s, rep)
End Function

' ���򣺽��ж�
Private Function ��������(ByVal s As String, ByVal pat As String) As Boolean
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True: rx.Global = False: rx.pattern = pat
    �������� = rx.TEST(s)
End Function

' ���ݵ�ͬĿ¼
Private Function ���ݵ�ǰ�ĵ�(ByVal doc As Document) As String
    On Error GoTo EH
    Dim baseName As String, ext As String, bak As String, folder As String, ts As String
    ts = Format(Now, "yyyymmdd_hhnnss")
    If Len(doc.name) > 0 Then
        baseName = doc.name
        If InStrRev(baseName, ".") > 0 Then
            ext = mid$(baseName, InStrRev(baseName, "."))
            baseName = Left$(doc.name, InStrRev(doc.name, ".") - 1)
        Else
            ext = ".docx"
        End If
    Else
        baseName = "δ�����ĵ�": ext = ".docx"
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

