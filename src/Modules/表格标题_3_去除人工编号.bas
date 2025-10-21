Attribute VB_Name = "������_3_ȥ���˹����"
Option Explicit
Private Const ����_��ӡ��� As Boolean = False


' ==========================================================
' ������ֹ��������ѭ�� + ���������壩
' ����
'   ��һ�����������ġ�Story�С����ף���Ԥ������ԡ�����ͷ�Ķ���
'   ������Step A���ӡ���ɾ������һ�������ַ���Ϊֹ���ɳԵ���ע�����Ŀ��Ʒ���
'   ������Step B�����ף��������塰��[��ѡ���ַ�]����[���]*[��ѡ-����]��
'   ���ģ�ÿ�ֽ������������>0������ѯ���Ƿ������=0 ����ʾ�������ϡ�
'   ���壩��������־��ʹ�������е� ProgressForm������Ĵ��壩
' ==========================================================

Public Sub ��������ֹ����_ʹ�ý��ȴ���1()
    Const ������ָ����ʽ As Boolean = True
    Const ������ʽ�� As String = "������"

    Dim passNo As Long
    Dim total As Long, touched As Long, skipped As Long
    Dim allTouched As Long
    Dim ans As VbMsgBoxResult
    Dim doc As Document: Set doc = ActiveDocument
    Dim useStyleFilter As Boolean
    Dim capStyle As Style

    ' ��һ����ʽ�����ж�
    useStyleFilter = ������ָ����ʽ
    If useStyleFilter Then
        On Error Resume Next
        Set capStyle = doc.Styles(������ʽ��)
        On Error GoTo 0
        If capStyle Is Nothing Then useStyleFilter = False
    End If

    ' ��������ʾ���ȴ��壨��ģʽ��
    With progressForm
        .caption = "�����ǰ׺���"
        .FrameProgress.width = 0
        .LabelPercentage.caption = "0%"
        .TextBoxStatus.text = "׼���С���" & vbCrLf
        .stopFlag = False
        .Show vbModeless
        DoEvents
    End With

'    Application.ScreenUpdating = False
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "�����ǰ׺�����ѭ����"
    On Error GoTo 0

    Do
        passNo = passNo + 1
        progressForm.TextBoxStatus.text = progressForm.TextBoxStatus.text & _
            "���� �� " & passNo & " �ֿ�ʼ ����" & vbCrLf

        ' ������ͳ�ƺ�ѡ���������ڼ�����ȣ�
        Dim cand As Long
        cand = ͳ�ƺ�ѡ����(useStyleFilter, ������ʽ��)

        If cand = 0 Then
            progressForm.UpdateProgressBar 200, "����û���ԡ�����ͷ�ĺ�ѡ���䡣"
            MsgBox "�����ǰ׺�����ϣ�δ���ֺ�ѡ����" & vbCrLf & _
                   "�ۼ��������" & allTouched & " ����", vbInformation
            Exit Do
        End If

        ' ���ģ�ִ�е�����������������־��
        ִ��һ����� cand, useStyleFilter, ������ʽ��, total, touched, skipped
        allTouched = allTouched + touched

        ' ���壩�ִ�С��
        Dim summary As String
        summary = "�� " & passNo & " ����ɣ���ѡ=" & total & "�������=" & touched & "��δ���=" & skipped
        progressForm.UpdateProgressBar 200, summary

        ' �������������� / ������ʾ
        If progressForm.stopFlag Then
            MsgBox "���ֶ���ֹ���ۼ��������" & allTouched & " ����", vbExclamation
            Exit Do
        End If

        If touched = 0 Then
            MsgBox "�����ǰ׺�����ϣ������޿�������" & vbCrLf & _
                   "�ۼ��������" & allTouched & " ����", vbInformation
            Exit Do
        Else
            ans = MsgBox("��������� " & touched & " ��ǰ׺��" & vbCrLf & _
                         "ǰ׺������δ��ȫ������Ƿ������һ�֣�", _
                         vbYesNo + vbQuestion, "���������")
            If ans = vbNo Then Exit Do
            ' ���ý�����
            progressForm.FrameProgress.width = 0
            progressForm.LabelPercentage.caption = "0%"
        End If

        ' ��ȫ�������⼫��ѭ��
        If passNo >= 5 Then
            ans = MsgBox("������ 5 �֣��Ƿ���Ҫ������", vbYesNo + vbExclamation)
            If ans = vbNo Then Exit Do
        End If
    Loop

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
    Application.ScreenUpdating = True

'    ' ��β�����ش���
'    On Error Resume Next
'    Unload progressForm
'    On Error GoTo 0
End Sub


' ----------------------------------------------------------
' ִ�е�����������������
' candTotal: ͳ�Ƶ��ĺ�ѡ���������ڽ��ȼ��㣩
' useStyleFilter / targetStyleName: �Ƿ�����������⡿��ʽ
' �����total/touched/skipped������ͳ�ƣ�
' ----------------------------------------------------------
Private Sub ִ��һ�����(ByVal candTotal As Long, _
                        ByVal useStyleFilter As Boolean, _
                        ByVal targetStyleName As String, _
                        ByRef total As Long, _
                        ByRef touched As Long, _
                        ByRef skipped As Long)

    Dim doc As Document: Set doc = ActiveDocument
    Dim p As Paragraph, r As Range
    Dim oldTxt As String, newTxt As String
    Dim processed As Long, progressPx As Long
    Dim examples As Long

    total = 0: touched = 0: skipped = 0

    For Each p In doc.Paragraphs
        If progressForm.stopFlag Then Exit For
        If p.Range.StoryType <> wdMainTextStory Then GoTo NextPara

'        ' ������1����ʽ���ˣ������ã�
'        If useStyleFilter Then
'            On Error Resume Next
'            If p.Range.Style.NameLocal <> targetStyleName Then GoTo NextPara
'            On Error GoTo 0
'        End If

        ' ������2����ѡ�ж������ף�ǿԤ��������ԡ�����ͷ
        oldTxt = ǿԤ����_����(p.Range.text)
        If Len(oldTxt) = 0 Or Left$(oldTxt, 1) <> "��" Then GoTo NextPara
        
        ' �������������ԣ�ԭ�� & ������
        If ����_��ӡ��� Then
            Debug.Print String(48, "-")
            Debug.Print "��ѡ#"; processed + 1; "/", candTotal
            Debug.Print "ԭ�ģ�"; Left$(p.Range.text, 120)
            Debug.Print "ԭ����㣺"; �г�ǰN���(p.Range.text, 60)
            Debug.Print "������"; oldTxt
            Debug.Print "������㣺"; �г�ǰN���(oldTxt, 60)
        End If

        total = total + 1
        processed = processed + 1

        ' ������3��Ŀ���ӷ�Χ��������β��ǣ�
        Set r = p.Range.Duplicate
        If r.Characters.Count > 1 Then r.MoveEnd wdCharacter, -1

        ' ������4��Step A���ӡ���ɾ�����׸����ġ�
        newTxt = ȥ�������ǰ׺_����һ������(r.text)

        ' ������5��Step B���� A δ�ı��ı��������򶵵�
        If newTxt = r.text Then
            newTxt = �����滻(newTxt, _
                    "^\s*��\s*[-���C��]?\s*\d+(?:\s*[\.����]\s*\d+)*\s*(?:[-���C��]\s*\d+)?(?:\s+\d+)?\s*", _
                    "")
            ' ���׺��������������ֱ���׸����ģ������ע���Ʒ�������
            newTxt = ȥ��ֱ���׸�����(newTxt)
        End If

        ' ������6����д & ��������־
        If newTxt <> r.text Then
            r.text = newTxt
            touched = touched + 1
            If examples < 6 Then
                progressForm.UpdateProgressBar ��ǰ��������(processed, candTotal), _
                    "���ǰ��" & Left$(oldTxt, 80) & vbCrLf & " �ĺ�" & Left$(ǿԤ����_����(newTxt), 80)
                examples = examples + 1
            Else
                progressForm.UpdateProgressBar ��ǰ��������(processed, candTotal), _
                    "�Ѵ���" & processed & "/" & candTotal
            End If
        Else
            skipped = skipped + 1
            progressForm.UpdateProgressBar ��ǰ��������(processed, candTotal), _
                "δ�����" & Left$(oldTxt, 80)
        End If

NextPara:
        DoEvents
    Next p
End Sub


Private Function ͳ�ƺ�ѡ����(ByVal useStyleFilter As Boolean, ByVal targetStyleName As String) As Long
    Dim doc As Document: Set doc = ActiveDocument
    Dim sty As Style
    Dim scope As Range, rng As Range
    Dim cnt As Long, nextPos As Long

    ' 1) �����������⡿��ʽ���������ڣ���ʾ���˳�
    On Error Resume Next
    Set sty = doc.Styles("������")
    On Error GoTo 0
    If sty Is Nothing Then
        MsgBox "����ִ�б������ʽƥ�䣡", vbExclamation
        ͳ�ƺ�ѡ���� = 0
        Exit Function
    End If

    ' 2) ����Χ������ǰ��ѡ���������ģ�ֻͳ��ѡ��������ͳ��ȫ������
    If Selection.Type <> wdSelectionIP And Selection.Range.StoryType = wdMainTextStory Then
        Set scope = Selection.Range.Duplicate
    Else
        Set scope = doc.StoryRanges(wdMainTextStory).Duplicate
    End If

    ' 3) �� Find ����ʽͳ�ƣ�ÿ�����к������ö�ĩ�������ظ�/����
    Set rng = scope.Duplicate
    With rng.Find
        .ClearFormatting
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = sty

        Do While .Execute
            cnt = cnt + 1
            nextPos = rng.Paragraphs(1).Range.End   ' �������ж���Ķ�β
            If nextPos >= scope.End Then Exit Do    ' ���ﷶΧĩβ�����
            rng.SetRange Start:=nextPos, End:=scope.End
        Loop
    End With

    ͳ�ƺ�ѡ���� = cnt
End Function



' ----------------------------------------------------------
' �������أ���Ĵ���200pxΪ���� �� ���� = 200 * done / total
' ----------------------------------------------------------
Private Function ��ǰ��������(ByVal done As Long, ByVal total As Long) As Long
    If total <= 0 Then
        ��ǰ�������� = 0
    Else
        ��ǰ�������� = CLng(200# * done / total)
        If ��ǰ�������� < 0 Then ��ǰ�������� = 0
        If ��ǰ�������� > 200 Then ��ǰ�������� = 200
    End If
End Function


' ================= ���Ĺ����� =================

'��һ��Step A�������ף�ǿԤ������ԡ�����ͷ���ѡ��������׸����ġ�֮��ȫ��ɾ��
Private Function ȥ�������ǰ׺_����һ������(ByVal s As String) As String
    Dim i As Long, ch As String, hit As Boolean
    s = ǿԤ����_����(s)
    If Len(s) = 0 Or Left$(s, 1) <> "��" Then
        ȥ�������ǰ׺_����һ������ = s
        Exit Function
    End If
    For i = 2 To Len(s)
        ch = mid$(s, i, 1)
        If �Ƿ������ַ�(ch) Then hit = True: Exit For
    Next i
    If hit Then
        ȥ�������ǰ׺_����һ������ = LTrim$(mid$(s, i))
    Else
        ȥ�������ǰ׺_����һ������ = s   ' û�ҵ����ģ����� Step B + ȥ�붵��
    End If
End Function

'������Step B ֮���ȥ�룺�Ѷ������С��հ�/���Ʒ�/���ַ�/���/���֡����뵽�׸�����
Private Function ȥ��ֱ���׸�����(ByVal s As String) As String
    Dim i As Long, ch As String
    s = ǿԤ����_����(s)
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If �Ƿ������ַ�(ch) Then ȥ��ֱ���׸����� = LTrim$(mid$(s, i)): Exit Function
    Next i
    ȥ��ֱ���׸����� = s
End Function


' ================= ��ϴ/�жϹ��� =================

'�����ߣ�����ǿԤ����ȥ��������ɼ������������� Trim
Private Function ǿԤ����_����(ByVal s As String) As String
    Dim i As Long, out As String, ch As String, cp As Long
    ' �����滻����β����Ԫ�������ȫ�ǿո�NBSP��Tab
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    s = Replace$(s, ChrW(&HA0), " ")
    s = Replace$(s, vbTab, " ")
    ' ���/�������
    s = Replace$(s, ChrW(&H200B), "")
    s = Replace$(s, ChrW(&H200C), "")
    s = Replace$(s, ChrW(&H200D), "")
    s = Replace$(s, ChrW(&HFEFF), "")
    s = Replace$(s, ChrW(&H200E), "")
    s = Replace$(s, ChrW(&H200F), "")
    s = Replace$(s, ChrW(&H202A), "")
    s = Replace$(s, ChrW(&H202B), "")
    s = Replace$(s, ChrW(&H202C), "")
    s = Replace$(s, ChrW(&H202D), "")
    s = Replace$(s, ChrW(&H202E), "")
    ' ��¼�ָ����ȿ��Ʒ����� U+001E��
    s = Replace$(s, ChrW(&H1E), "")
    ' ���գ��˳� <32 �Ŀ����ַ�
    out = ""
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + &H10000
        If cp >= 32 Then out = out & ch
    Next i
    ǿԤ����_���� = LTrim$(out)
End Function

'�����ߣ��ж��Ƿ����ģ����� AscW ������CJK ������ + ��չA��
Private Function �Ƿ������ַ�(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) = 0 Then �Ƿ������ַ� = False: Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + &H10000
    �Ƿ������ַ� = ((code >= &H4E00 And code <= &H9FFF) Or (code >= &H3400 And code <= &H4DBF))
End Function

'�����ߣ����򣺵����滻��ֻ�����һ�������࿿��ѭ�����֡���
Private Function �����滻(ByVal s As String, ByVal pat As String, Optional ByVal rep As String = "") As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Global = False
    rx.pattern = pat
    �����滻 = rx.Replace(s, rep)
End Function

'�����ߣ��г�ǰ N ���ַ��� Unicode ��㣨������ AscW ������������ "U+8868 U+002D ..."
Private Function �г�ǰN���(ByVal s As String, ByVal n As Long) As String
    Dim i As Long, out As String, cp As Long, ch As String
    Dim m As Long: m = IIf(Len(s) < n, Len(s), n)
    For i = 1 To m
        ch = mid$(s, i, 1)
        cp = AscW(ch)
        If cp < 0 Then cp = cp + &H10000
        out = out & "U+" & Right$("0000" & Hex$(cp), 4) & " "
    Next
    �г�ǰN��� = Trim$(out)
End Function

