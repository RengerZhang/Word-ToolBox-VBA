Attribute VB_Name = "������_5_Ԥ���"
'======================== ģ�飺������_5_Ԥ��� ========================
Option Explicit
Private Const BAR_MAX As Integer = 200   ' ProgressForm ����������

'�������ڸ������ȡ�Ĺ¶����б�ÿ�� [1]=���⼶��(1..4/0), [2]=�ı�, [3]=�����
Public gOrphanRows As Variant

Public Function GetOrphanRows() As Variant
    GetOrphanRows = gOrphanRows
End Function


'��һ����ڣ��Լ죨�������� �� �򿪴��壩
Public Sub �Լ�_��������ʽһ����()
    Dim doc As Document: Set doc = ActiveDocument
    Dim arrTableInfo As Variant
    Dim headStarts() As Long, headKeys() As String, headLvls() As Long, headNums() As String

    ' 1) �򿪽��ȴ������ֱ������� ProgressForm��
    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "���Ԥ��� - ��ʼ����"
    pf.Show vbModeless
    pf.TextBoxStatus.text = "���ڿ�ʼ������Ԥ���"
    DoEvents

    Application.ScreenUpdating = False

    ' 2) ����������30%��
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT

    ' 3) ���ɱ����飨��70%��
    arrTableInfo = BuildTableInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT

    Application.ScreenUpdating = True

    ' 4) ��������
    Dim frm As �Լ챨��
    Set frm = New �Լ챨��
    frm.LoadReportFromArray arrTableInfo
    frm.Show vbModeless

    ' 5) ���
    pf.UpdateProgressBar BAR_MAX, "��ɡ�"
    DoEvents
ABORT:
    Unload pf
End Sub
'��һ������ڣ��Լ죨ͼƬ���⣩�����߼�ͬ���ֻ�Ǹ���ͼƬ���� + �����е���ͼ��
Public Sub �Լ�_ͼƬ������ʽһ����()
    Dim doc As Document: Set doc = ActiveDocument
    Dim arrPicInfo As Variant
    Dim headStarts() As Long, headKeys() As String, headLvls() As Long, headNums() As String

    Dim pf As progressForm: Set pf = New progressForm
    pf.caption = "ͼƬԤ��� - ��ʼ����"
    pf.Show vbModeless
    pf.TextBoxStatus.text = "���ڿ�ʼͼƬ����Ԥ���"
    DoEvents

    Application.ScreenUpdating = False

    ' �� ������
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT

    ' �� ����ͼƬ����
    arrPicInfo = BuildImageInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT

    Application.ScreenUpdating = True

    ' �� �򿪴��壨ע�⴫ "ͼ"��
    Dim frm As �Լ챨��
    Set frm = New �Լ챨��
    frm.LoadReportFromArray arrPicInfo, "ͼ"
    frm.Show vbModeless

    pf.UpdateProgressBar BAR_MAX, "��ɡ�"
    DoEvents
ABORT:
    Unload pf
End Sub



'������������������������һ��ɨ�裬�����ȣ�0%��30%��
Private Sub BuildHeadingIndex(ByVal doc As Document, _
                              ByRef headStarts() As Long, ByRef headKeys() As String, _
                              ByRef headLvls() As Long, ByRef headNums() As String, _
                              ByRef pf As progressForm)

    Dim total As Long: total = doc.Paragraphs.Count
    Dim ts() As Long, tl() As Long, tn() As String, tk() As String
    ReDim ts(1 To total): ReDim tl(1 To total): ReDim tn(1 To total): ReDim tk(1 To total)

    Dim i As Long, cnt As Long
    Dim p As Paragraph

    ' �ȸ�������������Сռλ���������δ��ʼ��
    ReDim headStarts(1 To 1): headStarts(1) = 0
    ReDim headLvls(1 To 1): headLvls(1) = 0
    ReDim headNums(1 To 1): headNums(1) = ""
    ReDim headKeys(1 To 1): headKeys(1) = ""

    For Each p In doc.Paragraphs
        cnt = cnt + 1

        ' ��������ˢ�½��ȣ�ÿ200�Σ�
        If cnt Mod 200 = 0 Then
            Dim w As Integer
            w = CInt(BAR_MAX * 0.3 * cnt / IIf(total = 0, 1, total))
            pf.UpdateProgressBar w, "�� �������������� " & cnt & "/" & total
            DoEvents
            If pf.stopFlag Then Exit Sub
        End If

        ' ����ɸѡ����1~4
        Dim sty As String: sty = SafeStyleName(p.Range)
        Dim lvl As Long: lvl = HeadingLevelByStyle(sty)
        If lvl >= 1 And lvl <= 4 Then
            i = i + 1
            ts(i) = p.Range.Start
            tl(i) = lvl
            tn(i) = SafeListString(p.Range)
            tk(i) = NormalizeChapterKey(tn(i), lvl)
        End If
    Next

    If i = 0 Then Exit Sub

    ReDim headStarts(1 To i)
    ReDim headLvls(1 To i)
    ReDim headNums(1 To i)
    ReDim headKeys(1 To i)
    Dim k As Long
    For k = 1 To i
        headStarts(k) = ts(k)
        headLvls(k) = tl(k)
        headNums(k) = tn(k)
        headKeys(k) = tk(k)
    Next

    pf.UpdateProgressBar CInt(BAR_MAX * 0.3), "�� ��������������ɡ�"
    DoEvents
End Sub

'���������ɡ�����Ϣ���顱�������ȣ�30%��100%��
'   �ж��壨1-based��������ԭ 1~12 �в��䣬������
'   [13] ���¡��¶��Ρ�HTML�������κα�ǰ������ʽ=�������⡿��
'   ���¶��Ρ��Ĺ����°� headStarts/headKeys ��������Ͻ��ж�
'===================================================
Private Function BuildTableInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    Dim n As Long: n = doc.Tables.Count
    Dim arr As Variant

    '��һ���ޱ��ʱ�ȱ�Ƿ��ؿ����飬���Ի�����ռ��¶���
    If n = 0 Then
        BuildTableInfoArray = Empty
        gOrphanRows = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "δ��⵽�κα��"
        ' ��ʹ�ޱ�������Ի��ռ��¶��Σ������� Collect��
    End If

    '�������ռ�����ʽ=�����⡱�����жΣ�λ��/�ı�/�����¼�
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    CollectAllCaptionParas doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt

    '����������������飬����¼����Ч�ı�ǰ����Ρ�
    If n > 0 Then ReDim arr(1 To n, 1 To 12)
    Dim dictSeq As Object:        Set dictSeq = CreateObject("Scripting.Dictionary")      ' �¼����������
    Dim dictValidCap As Object:   Set dictValidCap = CreateObject("Scripting.Dictionary") ' ��ǰ���������True

    Dim i As Long, tbl As Table
    Dim baseW As Integer: baseW = 60

    For i = 1 To n
        If Not pf Is Nothing Then
            If (i Mod 10 = 0) Or (i = n) Then
                pf.UpdateProgressBar baseW + CInt((140# / IIf(n = 0, 1, n)) * i), "�� ���ɱ���Ϣ�� " & i & "/" & n
            End If
        End If

        Set tbl = doc.Tables(i)
        Dim tStart As Long: tStart = tbl.Range.Start

        ' 1) �ͽ���һ�ǿն�
        Dim pstart As Long: pstart = PrevNonEmptyParaStart_ByStart(doc, tStart)
        Dim pTxt As String, pSty As String
        If pstart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, pstart))
            pSty = SafeStyleNameByStart(pstart)
        Else
            pTxt = "": pSty = ""
        End If

        ' ��¼����Ч���⡱��ȷʵλ�ڱ�ǰ�ı����⣩
        If (pstart > 0) And (pSty = "������") Then
            dictValidCap(CStr(pstart)) = True
        End If

        ' 2) ������⣨�Ͻ���֣�
        Dim idx As Long: idx = UpperBoundByStart(headStarts, tStart)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If

        ' 3) �����������
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "��" & key & "-" & CStr(seq)

        ' 4) �Ƿ����С������⡿
        Dim isCap As Boolean: isCap = (pSty = "������")

        ' 5) д�����飨����ԭ 1~12 �ж��岻�䣩
        arr(i, 1) = i
        arr(i, 2) = tStart
        arr(i, 3) = pstart
        arr(i, 4) = pTxt
        arr(i, 5) = pSty
        arr(i, 6) = isCap
        arr(i, 7) = lvl
        arr(i, 8) = hText
        arr(i, 9) = num
        arr(i, 10) = key
        arr(i, 11) = seq
        arr(i, 12) = label
    Next

    '���ģ��������ɡ��¶��Ρ���ά���飺ÿ�� [lvl, text, startPos]
    Dim orCnt As Long: orCnt = 0
    Dim orArr As Variant

    If capCnt > 0 Then
        Dim j As Long
        For j = 1 To capCnt
            Dim cs As Long: cs = capStarts(j)
            If cs <= 0 Then GoTo NEXT_J

            If Not dictValidCap.exists(CStr(cs)) Then
                Dim lvl2 As Long
                Dim idx2 As Long: idx2 = UpperBoundByStart(headStarts, cs)
                If idx2 >= 1 Then
                    lvl2 = headLvls(idx2)
                Else
                    lvl2 = 0
                End If

                orCnt = orCnt + 1
                ' �����������һά�ɱ䡱�ķ�ʽ���ݣ�orArr(3, ��)
                If orCnt = 1 Then
                    ReDim orArr(1 To 3, 1 To 1)
                Else
                    ReDim Preserve orArr(1 To 3, 1 To orCnt)   ' ֻ�ܸ����һά
                End If
                ' д�뵱ǰ�У�ע���±����ı䣩
                orArr(1, orCnt) = lvl2
                orArr(2, orCnt) = capTexts(j)
                orArr(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If        ' ������ ������ȱʧ�� End If�����ڱպ� If capCnt > 0 Then

    '���壩���������ʹ�õ� gOrphanRows��ת�� �С�3 ��״��
    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, k As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For k = 1 To orCnt
            outArr(k, 1) = orArr(1, k)   ' lvl
            outArr(k, 2) = orArr(2, k)   ' text
            outArr(k, 3) = orArr(3, k)   ' start
        Next
        gOrphanRows = outArr
    End If

    BuildTableInfoArray = arr
End Function





'�������ģ����ߣ��ռ��ĵ���������ʽ=�������⡿�Ķ��䣨λ�á��ı��������¼���
Private Sub CollectAllCaptionParas(ByVal doc As Document, _
                                   ByRef headStarts() As Long, ByRef headKeys() As String, _
                                   ByRef capStarts() As Long, ByRef capTexts() As String, _
                                   ByRef capKeys() As String, ByRef capCnt As Long, _
                                   Optional ByVal styleName As String = "������")

    Dim total As Long: total = doc.Paragraphs.Count
    ReDim capStarts(1 To total)
    ReDim capTexts(1 To total)
    ReDim capKeys(1 To total)
    Dim p As Paragraph, i As Long

    For Each p In doc.Paragraphs
        If SafeStyleName(p.Range) = styleName Then
            i = i + 1
            capStarts(i) = p.Range.Start
            capTexts(i) = TrimVisible(FirstParaTextAtStart(doc, p.Range.Start))
            Dim idx As Long: idx = UpperBoundByStart(headStarts, p.Range.Start)
            If idx >= 1 Then
                capKeys(i) = headKeys(idx)
            Else
                capKeys(i) = "0"
            End If
        End If
    Next
    If i = 0 Then
        ReDim capStarts(1 To 1): capStarts(1) = 0
        ReDim capTexts(1 To 1):  capTexts(1) = ""
        ReDim capKeys(1 To 1):   capKeys(1) = "0"
        capCnt = 0
    Else
        ReDim Preserve capStarts(1 To i)
        ReDim Preserve capTexts(1 To i)
        ReDim Preserve capKeys(1 To i)
        capCnt = i
    End If
End Sub

'============================== ���ߺ�������ǰ��һ�£� ==============================
Private Function SafeStyleName(ByVal r As Range) As String
    On Error Resume Next
    SafeStyleName = r.Style.nameLocal
    If Err.Number <> 0 Then SafeStyleName = ""
    On Error GoTo 0
End Function

Private Function SafeStyleNameByStart(ByVal pstart As Long) As String
    On Error Resume Next
    SafeStyleNameByStart = ActiveDocument.Range(pstart, pstart).Paragraphs(1).Range.Style.nameLocal
    If Err.Number <> 0 Then SafeStyleNameByStart = ""
    On Error GoTo 0
End Function

Private Function HeadingLevelByStyle(ByVal sty As String) As Long
    Select Case sty
        Case "���� 1", "����1": HeadingLevelByStyle = 1
        Case "���� 2", "����2": HeadingLevelByStyle = 2
        Case "���� 3", "����3": HeadingLevelByStyle = 3
        Case "���� 4", "����4": HeadingLevelByStyle = 4
        Case Else: HeadingLevelByStyle = 0
    End Select
End Function

Private Function SafeListString(ByVal r As Range) As String
    On Error Resume Next
    SafeListString = r.ListFormat.ListString
    If Err.Number <> 0 Then SafeListString = ""
    On Error GoTo 0
End Function

Private Function NormalizeChapterKey(ByVal listStr As String, ByVal lvl As Long) As String
    Dim s As String: s = Trim$(listStr)
    If s <> "" Then
        NormalizeChapterKey = s
    ElseIf lvl > 0 Then
        NormalizeChapterKey = "L" & CStr(lvl)
    Else
        NormalizeChapterKey = ""
    End If
End Function

Private Function PrevNonEmptyParaStart_ByStart(ByVal doc As Document, ByVal atStart As Long) As Long
    Dim prgs As Paragraphs
    On Error Resume Next
    Set prgs = doc.Range(0, atStart).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then
        PrevNonEmptyParaStart_ByStart = 0: Exit Function
    End If
    Dim p As Paragraph: Set p = prgs(prgs.Count)
    On Error GoTo 0
    Do While Not p Is Nothing
        Dim s As String: s = TrimVisible(p.Range.text)
        If s <> "" Then PrevNonEmptyParaStart_ByStart = p.Range.Start: Exit Function
        Set p = p.Previous
    Loop
    PrevNonEmptyParaStart_ByStart = 0
End Function

Private Function FirstParaTextAtStart(ByVal doc As Document, ByVal pstart As Long) As String
    If pstart <= 0 Then Exit Function
    Dim r As Range
    Set r = doc.Range(Start:=pstart, End:=doc.Range(pstart, pstart).Paragraphs(1).Range.End)
    FirstParaTextAtStart = TrimVisible(r.text)
End Function

Private Function TrimVisible(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr(7), "")
    s = Replace$(s, ChrW(&H3000), " ")
    TrimVisible = Trim$(s)
End Function

Private Function UpperBoundByStart(ByRef starts() As Long, ByVal atStart As Long) As Long
    If (Not Not starts) = 0 Then Exit Function
    Dim lo As Long, hi As Long, mid As Long, ans As Long
    lo = 1
    hi = UBound(starts)
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If starts(mid) < atStart Then ans = mid: lo = mid + 1 Else hi = mid - 1
    Loop
    UpperBoundByStart = ans
End Function

'������������HTML ת�壺����ģ���ڲ�ʹ��
Private Function HtmlEncode(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function


'�����������ɡ�ͼƬ��Ϣ���顱�������ȣ�30%��100%��
'   �������� 1..12 �У�
'   [1]=��� [2]=ͼƬ��� [3]=�������㣨�·���һ���ǿնΣ�
'   [4]=������ı� [5]=�������ʽ [6]=�Ƿ�=��ͼƬ���⡿
'   [7..11]=ͬ��������⼶��/�ı�/���/�¼�/�������
'   [12]=��ͼ<�¼�>-<���>��
Private Function BuildImageInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    Dim nInline As Long: nInline = doc.InlineShapes.Count
    Dim nShape As Long:  nShape = CountPictureShapes_InModule(doc)
    Dim total As Long:   total = nInline + nShape
    Dim baseW As Integer: baseW = 60

    Dim arr As Variant
    If total = 0 Then
        BuildImageInfoArray = Empty
        gOrphanRows = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "δ��⵽�κ�ͼƬ��"
        Exit Function
    End If

    ' �����ռ�����ͼƬ�ġ��ĵ�λ����㡱�롰�������㡱
    Dim pos() As Long, pstart() As Long, cnt As Long
    ReDim pos(1 To total): ReDim pstart(1 To total)

    Dim i As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        i = i + 1
        pos(i) = ils.Range.Start
        pstart(i) = NextNonEmptyParaStart_ByStart(doc, ils.Range.Start)
    Next

    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape_InModule(s) Then
            i = i + 1
            pos(i) = s.anchor.Start
            pstart(i) = NextNonEmptyParaStart_ByStart(doc, s.anchor.Start)
        End If
    Next
    cnt = i

    ' ������λ�����򣨱�֤�ĵ��Ķ�˳��
    Call SelectionSortByPos(pos, pstart, cnt)

    ' ����������
    ReDim arr(1 To cnt, 1 To 12)

    Dim dictSeq As Object:      Set dictSeq = CreateObject("Scripting.Dictionary")  ' �¼����������
    Dim dictValidCap As Object: Set dictValidCap = CreateObject("Scripting.Dictionary") ' ��Ч���������True

    Dim k As Long
    For k = 1 To cnt
        If Not pf Is Nothing Then
            If (k Mod 10 = 0) Or (k = cnt) Then
                pf.UpdateProgressBar baseW + CInt((140# / IIf(cnt = 0, 1, cnt)) * k), "�� ����ͼƬ��Ϣ�� " & k & "/" & cnt
            End If
        End If

        Dim atPos As Long: atPos = pos(k)
        Dim capStart As Long: capStart = pstart(k)

        ' ������ı�/��ʽ
        Dim pTxt As String, pSty As String
        If capStart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, capStart))
            pSty = SafeStyleNameByStart(capStart)
        Else
            pTxt = "": pSty = ""
        End If
        Dim isCap As Boolean: isCap = (pSty = "ͼƬ����")
        If isCap And capStart > 0 Then dictValidCap(CStr(capStart)) = True

        ' ������⣨�Ͻ磩
        Dim idx As Long: idx = UpperBoundByStart(headStarts, atPos)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If

        ' ��������롰ͼ�š�
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "ͼ" & key & "-" & CStr(seq)

        ' д�루���� 1..12 �У�
        arr(k, 1) = k
        arr(k, 2) = atPos
        arr(k, 3) = capStart
        arr(k, 4) = pTxt
        arr(k, 5) = pSty
        arr(k, 6) = isCap
        arr(k, 7) = lvl
        arr(k, 8) = hText
        arr(k, 9) = num
        arr(k, 10) = key
        arr(k, 11) = seq
        arr(k, 12) = label
    Next

    ' �������ɡ��¶��Ρ�����ʽ=ͼƬ���⣬��δ���κ�ͼƬ���У�
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    CollectAllCaptionParas doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt, "ͼƬ����"

    Dim orCnt As Long: orCnt = 0
    Dim orArr As Variant
    If capCnt > 0 Then
        Dim j As Long
        For j = 1 To capCnt
            Dim cs As Long: cs = capStarts(j)
            If cs <= 0 Then GoTo NEXT_J
            If Not dictValidCap.exists(CStr(cs)) Then
                Dim lvl2 As Long
                Dim idx2 As Long: idx2 = UpperBoundByStart(headStarts, cs)
                If idx2 >= 1 Then lvl2 = headLvls(idx2) Else lvl2 = 0
                orCnt = orCnt + 1
                If orCnt = 1 Then
                    ReDim orArr(1 To 3, 1 To 1)
                Else
                    ReDim Preserve orArr(1 To 3, 1 To orCnt)
                End If
                orArr(1, orCnt) = lvl2
                orArr(2, orCnt) = capTexts(j)
                orArr(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If

    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, t As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For t = 1 To orCnt
            outArr(t, 1) = orArr(1, t)
            outArr(t, 2) = orArr(2, t)
            outArr(t, 3) = orArr(3, t)
        Next
        gOrphanRows = outArr
    End If

    BuildImageInfoArray = arr
End Function

' ������������ǰλ�ú��·���һ���ǿնΡ������
Private Function NextNonEmptyParaStart_ByStart(ByVal doc As Document, ByVal atStart As Long) As Long
    Dim prgs As Paragraphs
    On Error Resume Next
    Set prgs = doc.Range(atStart, doc.content.End).Paragraphs
    If prgs Is Nothing Or prgs.Count = 0 Then Exit Function
    Dim p As Paragraph: Set p = prgs(1).Next    ' �ӡ���һ�����䡱��ʼ
    On Error GoTo 0
    Do While Not p Is Nothing
        Dim s As String: s = TrimVisible(p.Range.text)
        If s <> "" Then NextNonEmptyParaStart_ByStart = p.Range.Start: Exit Function
        Set p = p.Next
    Loop
    NextNonEmptyParaStart_ByStart = 0
End Function

' �����������ж� Shape �Ƿ�ΪͼƬ
Private Function IsPictureShape_InModule(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsPictureShape_InModule = (s.Type = msoPicture Or s.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

' ����������ͳ��ͼƬ�͸��� Shape
Private Function CountPictureShapes_InModule(ByVal doc As Document) As Long
    Dim sh As Shape, n As Long
    For Each sh In doc.Shapes
        If IsPictureShape_InModule(sh) Then n = n + 1
    Next
    CountPictureShapes_InModule = n
End Function

' �������������ĵ�λ���������У�ԭ�ؽ��� pos��pstart �������飩
Private Sub SelectionSortByPos(ByRef pos() As Long, ByRef pstart() As Long, ByVal n As Long)
    Dim i As Long, j As Long, imin As Long, tp As Long, ts As Long
    For i = 1 To n - 1
        imin = i
        For j = i + 1 To n
            If pos(j) < pos(imin) Then imin = j
        Next
        If imin <> i Then
            tp = pos(i): pos(i) = pos(imin): pos(imin) = tp
            ts = pstart(i): pstart(i) = pstart(imin): pstart(imin) = ts
        End If
    Next
End Sub


