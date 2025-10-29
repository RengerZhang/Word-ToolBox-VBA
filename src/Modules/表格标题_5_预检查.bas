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
    
    
    Dim t As Double: t = t0()                     ' �� ��ʱ���
    PF_Log pf, "�� �ر���Ļˢ��..."
    Application.ScreenUpdating = False
    PF_Tick pf, t, "�ر���Ļˢ��", 10              ' �� ���

    ' 2) ����������30%��
    PF_Log pf, "�� ��������������ɨ�����Ķ��䣩..."
    BuildHeadingIndex doc, headStarts, headKeys, headLvls, headNums, pf
    If pf.stopFlag Then GoTo ABORT
    PF_Tick pf, t, "������������", 70              ' �� ���

    ' 3) ���ɱ����飨��70%��
    PF_Log pf, "�� ��������Ϣ���飨ɨ�������ǰ�Ρ��¶��Σ�..."
    arrTableInfo = BuildTableInfoArray(doc, headStarts, headKeys, headLvls, headNums, pf)
    If pf.stopFlag Then GoTo ABORT
    PF_Tick pf, t, "��������Ϣ����", 160            ' �� ���
    
    
    PF_Log pf, "�� ��Ⱦ���浽 WebBrowser..."
    Application.ScreenUpdating = True
    PF_Tick pf, t, "����Ļˢ��", 170

    ' 4) ��������
    Dim frm As �Լ챨��
    Set frm = New �Լ챨��
    frm.LoadReportFromArray arrTableInfo
    frm.Show vbModeless
    PF_Tick pf, t, "����/д�� HTML", 200

    ' 5) ���
    pf.UpdateProgressBar BAR_MAX, "��ɡ�"
    DoEvents
ABORT:
    'Unload pf
    
    
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
'=========================================================
' �������Ԥ��������� + �¶����嵥
' ���أ�arr(1..N, 1..12)
'   1=������  2=��Range.Start  3=��ǰ�����  4=��ǰ���ı�
'   5=��ǰ����ʽ  6=�Ƿ����С������⡱  7=������⼶��
'   8=��������ı�  9=��������Ŵ�  10=�¼�  11=�������
'   12=��ţ��硰��3.1-2����
' ͬʱ�����gOrphanRows(1..M,1..3) �� [lvl, text, start]
'=========================================================
Private Function BuildTableInfoArray(ByVal doc As Document, _
                                     ByRef headStarts() As Long, ByRef headKeys() As String, _
                                     ByRef headLvls() As Long, ByRef headNums() As String, _
                                     Optional ByRef pf As progressForm) As Variant
    '���㣩��������׼��
    Dim N As Long: N = doc.Tables.Count
    Dim arr As Variant
    Dim basePx As Long: basePx = 70           ' ���׶ν��Ȼ��ߣ�����ڶ�Ӧ��
    Dim i As Long, tbl As Table
    Dim t0 As Double                          ' ������ʱ��
    Dim sumPrev As Double, sumHead As Double, sumWrite As Double ' �����ۼ�

    '��һ���ޱ��ʱ�������ռ��¶��Σ��������鷵�� Empty
    If N = 0 Then
        BuildTableInfoArray = Empty
        If Not pf Is Nothing Then pf.UpdateProgressBar 60, "δ��⵽�κα���Խ��ռ��¶��Σ���"
    Else
        ReDim arr(1 To N, 1 To 12)
    End If

    '�������ռ����С���ʽ=�����⡱�ĶΣ��������¶��β�ã�
    Dim capStarts() As Long, capTexts() As String, capKeys() As String, capCnt As Long
    t0 = Timer
    Call CollectAllCaptionParas(doc, headStarts, headKeys, capStarts, capTexts, capKeys, capCnt, pf)
    If Not pf Is Nothing Then PF_StepWarn pf, t0, "��-Ԥȡ���С������⡯��", 1, IIf(N = 0, 1, N), 0.08, 8, basePx, 10

    '��������������У�����¼����Ч�ı�ǰ����Ρ����ֵ䣩
    Dim dictSeq As Object:      Set dictSeq = CreateObject("Scripting.Dictionary")   ' �¼����������
    Dim dictValidCap As Object: Set dictValidCap = CreateObject("Scripting.Dictionary") ' ��Ч��������True

    For i = 1 To N
        ' �������ȣ��ܽ�������70��150��
        If Not pf Is Nothing Then
            If (i Mod 10 = 0) Or (i = N) Then
                pf.UpdateProgressBar basePx + CInt((80# / IIf(N = 0, 1, N)) * i), "�� ɨ���� " & i & "/" & N
            End If
        End If

        Set tbl = doc.Tables(i)
        Dim tStart As Long: tStart = tbl.Range.Start

        '��1���ͽ�����һ�ǿնΡ� �� ��Ϊ����κ�ѡ
        t0 = Timer
        Dim pStart As Long: pStart = PrevNonEmptyParaStart_ByStart(doc, tStart)
        sumPrev = sumPrev + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "�� ����һ�ǿն�", i, N, 0.12, 10, basePx, 80

        '��2����ȡ�öε��ı�����ʽ
        t0 = Timer
        Dim pTxt As String, pSty As String
        If pStart > 0 Then
            pTxt = TrimVisible(FirstParaTextAtStart(doc, pStart))
            pSty = SafeStyleNameByStart(pStart)
        Else
            pTxt = "": pSty = ""
        End If
        sumWrite = sumWrite + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "�� ��ȡ������ı�/��ʽ", i, N, 0.08, 12, basePx, 80

        ' �������С������⡱���Ϊ��Ч���⣨���ں���¶��β��
        If (pStart > 0) And (pSty = "������") Then
            dictValidCap(CStr(pStart)) = True
        End If

        '��3����λ��������⡱����Ԥ������ head* ��������Ͻ磩
        t0 = Timer
        Dim idx As Long: idx = UpperBoundByStart(headStarts, tStart)
        Dim key As String, lvl As Long, num As String, hText As String
        If idx >= 1 Then
            key = headKeys(idx):  lvl = headLvls(idx)
            num = headNums(idx):  hText = FirstParaTextAtStart(doc, headStarts(idx))
        Else
            key = "0": lvl = 0: num = "": hText = ""
        End If
        sumHead = sumHead + (Timer - t0)
        If Not pf Is Nothing Then PF_StepWarn pf, t0, "�� ���ֶ�λ�������", i, N, 0.1, 10, basePx, 80

        '��4����������롰��š����㣨�¼��� headKeys��
        Dim seq As Long
        If Not dictSeq.exists(key) Then
            dictSeq.Add key, 1: seq = 1
        Else
            dictSeq(key) = dictSeq(key) + 1: seq = dictSeq(key)
        End If
        Dim label As String: label = "��" & key & "-" & CStr(seq)

        '��5���Ƿ����С������⡿��ʽ
        Dim isCap As Boolean: isCap = (pSty = "������")

        '��6��д�������飨����ԭ 1~12 �ж��岻�䣩
        arr(i, 1) = i
        arr(i, 2) = tStart
        arr(i, 3) = pStart
        arr(i, 4) = pTxt
        arr(i, 5) = pSty
        arr(i, 6) = isCap
        arr(i, 7) = lvl
        arr(i, 8) = hText
        arr(i, 9) = num
        arr(i, 10) = key
        arr(i, 11) = seq
        arr(i, 12) = label
    Next i

    '���ģ����ɡ��¶��Ρ���ά���飨���˱�������ʽ���������κα�ǰ��
    '     �������á����������ݡ��Ļ����������� O(n^2) �� ReDim Preserve
    Dim tOrp As Double: tOrp = Timer
    If Not pf Is Nothing Then PF_Log pf, "��-2 �ռ��¶��Ρ�"

    Dim orCnt As Long: orCnt = 0
    Dim cap As Long, orBuf As Variant
    If capCnt > 0 Then
        cap = IIf(capCnt \ 4 < 8, 8, capCnt \ 4)   ' ��ʼ������capCnt �� 1/4������ 8
        ReDim orBuf(1 To 3, 1 To cap)
        Dim j As Long
        For j = 1 To capCnt
            Dim cs As Long: cs = capStarts(j)
            If cs <= 0 Then GoTo NEXT_J

            If Not dictValidCap.exists(CStr(cs)) Then
                ' �������ڱ�ǰ����Ϊ�¶���
                Dim lvl2 As Long, idx2 As Long
                idx2 = UpperBoundByStart(headStarts, cs)
                If idx2 >= 1 Then lvl2 = headLvls(idx2) Else lvl2 = 0

                ' ���ݣ�2 ������д��
                orCnt = orCnt + 1
                If orCnt > cap Then
                    cap = cap * 2
                    ReDim Preserve orBuf(1 To 3, 1 To cap)
                End If
                orBuf(1, orCnt) = lvl2
                orBuf(2, orCnt) = capTexts(j)
                orBuf(3, orCnt) = cs
            End If
NEXT_J:
        Next j
    End If

    '���壩���������ʹ�õ� gOrphanRows��ѹ���� �С�3��
    If orCnt = 0 Then
        gOrphanRows = Empty
    Else
        Dim outArr As Variant, k As Long
        ReDim outArr(1 To orCnt, 1 To 3)
        For k = 1 To orCnt
            outArr(k, 1) = orBuf(1, k)   ' lvl
            outArr(k, 2) = orBuf(2, k)   ' text
            outArr(k, 3) = orBuf(3, k)   ' start
        Next k
        gOrphanRows = outArr
    End If

    If Not pf Is Nothing Then
        pf.UpdateProgressBar 160, "��-2 �¶�����ɣ���ʱ " & Format(Timer - tOrp, "0.000") & " s"
        pf.UpdateProgressBar 160, "�� �ܽ᣺prev=" & Format(sumPrev, "0.000") & "s��head=" & _
                                   Format(sumHead, "0.000") & "s��read/write=" & Format(sumWrite, "0.000") & "s"
    End If

    '����������������
    BuildTableInfoArray = arr
End Function






'�������ģ����ߣ��ռ��ĵ���������ʽ=�������⡿�Ķ��䣨λ�á��ı��������¼���
'Private Sub CollectAllCaptionParas(ByVal doc As Document, _
'                                   'ByRef headStarts() As Long, ByRef headKeys() As String, _
'                                   'ByRef capStarts() As Long, ByRef capTexts() As String, _
'                                   'ByRef capKeys() As String, ByRef capCnt As Long, _
'                                   'Optional ByVal styleName As String = "������")

    'Dim total As Long: total = doc.Paragraphs.Count
    'ReDim capStarts(1 To total)
    'ReDim capTexts(1 To total)
    'ReDim capKeys(1 To total)
    'Dim p As Paragraph, i As Long

    'For Each p In doc.Paragraphs
        'If SafeStyleName(p.Range) = styleName Then
            'i = i + 1
            'capStarts(i) = p.Range.Start
            'capTexts(i) = TrimVisible(FirstParaTextAtStart(doc, p.Range.Start))
            'Dim idx As Long: idx = UpperBoundByStart(headStarts, p.Range.Start)
            'If idx >= 1 Then
                'capKeys(i) = headKeys(idx)
            'Else
                'capKeys(i) = "0"
            'End If
        'End If
    'Next
    'If i = 0 Then
        'ReDim capStarts(1 To 1): capStarts(1) = 0
        'ReDim capTexts(1 To 1):  capTexts(1) = ""
        'ReDim capKeys(1 To 1):   capKeys(1) = "0"
        'capCnt = 0
    'Else
        'ReDim Preserve capStarts(1 To i)
        'ReDim Preserve capTexts(1 To i)
        'ReDim Preserve capKeys(1 To i)
        'capCnt = i
    'End If
'End Sub

'=========================================================
' Ԥȡ���С������⡱�Σ�����ϸ���ȷ�����
' �����
'   capStarts()  ÿ�����жε� Range.Start
'   capTexts()   ÿ�����жεĿɼ��ı����� Trim��
'   capKeys()    ÿ�����ж��������¼����� headKeys/UpperBoundByStart ӳ�䣩
'   capCnt       ��������
' ˵����
'   - ���ı�ԭ��������壬���������ȷ���������ˢ�������� + ������
'   - ����������ӳ�䣺62 �� 70���� BuildTableInfoArray �еĽ׶����������
'=========================================================
Private Sub CollectAllCaptionParas(ByVal doc As Document, _
                                   ByRef headStarts() As Long, ByRef headKeys() As String, _
                                   ByRef capStarts() As Long, ByRef capTexts() As String, _
                                   ByRef capKeys() As String, ByRef capCnt As Long, _
                                   Optional ByRef pf As progressForm)

    Dim totalP As Long: totalP = doc.Paragraphs.Count
    Dim cap As Long, hit As Long
    Dim i As Long, tStep As Double, tLast As Double
    Dim px As Long, basePx As Long: basePx = 62
    Dim spanPx As Long: spanPx = 8         ' �� 62..70

    '��һ����ʼ�����飨�� 1/64 �ĵ�����Ԥ�������� 64��
    cap = IIf(totalP \ 64 > 64, totalP \ 64, 64)
    ReDim capStarts(1 To cap)
    ReDim capTexts(1 To cap)
    ReDim capKeys(1 To cap)
    capCnt = 0

    '������������ʽ����ʧ�����˻�Ϊ���ƱȽ�
    Dim styCap As Style, safeByName As Boolean
    On Error Resume Next
    Set styCap = doc.Styles("������")
    On Error GoTo 0
    safeByName = (styCap Is Nothing)

    '������������ʾ
    If Not pf Is Nothing Then
        pf.UpdateProgressBar basePx, "��-Ԥȡ�������⡯�Σ���ʼɨ�裨�� " & totalP & " �Σ���"
        DoEvents
    End If

    tLast = Timer
    For i = 1 To totalP
        tStep = Timer

        Dim p As Paragraph
        Set p = doc.Paragraphs(i)

        ' ����������
        If p.Range.StoryType = wdMainTextStory Then
            ' �����ж������ȶ���Ƚϣ�������ƶ���
            Dim isCap As Boolean
            If Not safeByName Then
                On Error Resume Next
                isCap = (p.Range.Style Is styCap)
                If Err.Number <> 0 Then
                    Err.Clear
                    isCap = (CStr(p.Range.Style) = "������")
                End If
                On Error GoTo 0
            Else
                isCap = (CStr(p.Range.Style) = "������")
            End If

            If isCap Then
                ' ���У����ݣ�������
                hit = hit + 1
                If hit > cap Then
                    cap = cap * 2
                    ReDim Preserve capStarts(1 To cap)
                    ReDim Preserve capTexts(1 To cap)
                    ReDim Preserve capKeys(1 To cap)
                End If

                Dim st As Long: st = p.Range.Start
                capStarts(hit) = st
                capTexts(hit) = TrimVisible(FirstParaTextAtStart(doc, st))

                ' �����¼����ô���� headStarts/headKeys �����Ͻ�
                Dim idx As Long: idx = UpperBoundByStart(headStarts, st)
                If idx >= 1 Then
                    capKeys(hit) = headKeys(idx)
                Else
                    capKeys(hit) = "0"
                End If
            End If
        End If

        ' �������ģ����� + ����ʽ���ȷ�����ÿ 1000 �� �� ÿ ~0.6s �ر�һ��
        If Not pf Is Nothing Then
            If (i Mod 1000 = 0) Or (Timer - tLast >= 0.6) Or (i = totalP) Then
                tLast = Timer
                px = basePx + CLng(spanPx * i / IIf(totalP = 0, 1, totalP))
                pf.UpdateProgressBar px, _
                    "��-Ԥȡ���ȣ�" & i & "/" & totalP & "  | ���� " & hit & " ��"
                DoEvents
            End If
        End If
    Next i

    '���壩��β��ѹ�������д�С
    If hit = 0 Then
        Erase capStarts: Erase capTexts: Erase capKeys
        capCnt = 0
        If Not pf Is Nothing Then
            pf.UpdateProgressBar basePx + spanPx, "��-Ԥȡ��ɣ�δ�����κΡ������⡯�Ρ�"
        End If
        Exit Sub
    End If

    If hit < cap Then
        ReDim Preserve capStarts(1 To hit)
        ReDim Preserve capTexts(1 To hit)
        ReDim Preserve capKeys(1 To hit)
    End If
    capCnt = hit

    If Not pf Is Nothing Then
        pf.UpdateProgressBar basePx + spanPx, _
            "��-Ԥȡ��ɣ������� " & capCnt & " �Ρ�"
        DoEvents
    End If
End Sub


'============================== ���ߺ�������ǰ��һ�£� ==============================
Private Function SafeStyleName(ByVal r As Range) As String
    On Error Resume Next
    SafeStyleName = r.Style.NameLocal
    If Err.Number <> 0 Then SafeStyleName = ""
    On Error GoTo 0
End Function

Private Function SafeStyleNameByStart(ByVal pStart As Long) As String
    On Error Resume Next
    SafeStyleNameByStart = ActiveDocument.Range(pStart, pStart).Paragraphs(1).Range.Style.NameLocal
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

Private Function FirstParaTextAtStart(ByVal doc As Document, ByVal pStart As Long) As String
    If pStart <= 0 Then Exit Function
    Dim r As Range
    Set r = doc.Range(Start:=pStart, End:=doc.Range(pStart, pStart).Paragraphs(1).Range.End)
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
    Dim pos() As Long, pStart() As Long, cnt As Long
    ReDim pos(1 To total): ReDim pStart(1 To total)

    Dim i As Long, ils As InlineShape
    For Each ils In doc.InlineShapes
        i = i + 1
        pos(i) = ils.Range.Start
        pStart(i) = NextNonEmptyParaStart_ByStart(doc, ils.Range.Start)
    Next

    Dim s As Shape
    For Each s In doc.Shapes
        If IsPictureShape_InModule(s) Then
            i = i + 1
            pos(i) = s.anchor.Start
            pStart(i) = NextNonEmptyParaStart_ByStart(doc, s.anchor.Start)
        End If
    Next
    cnt = i

    ' ������λ�����򣨱�֤�ĵ��Ķ�˳��
    Call SelectionSortByPos(pos, pStart, cnt)

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
        Dim capStart As Long: capStart = pStart(k)

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
    Dim sh As Shape, N As Long
    For Each sh In doc.Shapes
        If IsPictureShape_InModule(sh) Then N = N + 1
    Next
    CountPictureShapes_InModule = N
End Function

' �������������ĵ�λ���������У�ԭ�ؽ��� pos��pstart �������飩
Private Sub SelectionSortByPos(ByRef pos() As Long, ByRef pStart() As Long, ByVal N As Long)
    Dim i As Long, j As Long, imin As Long, tp As Long, ts As Long
    For i = 1 To N - 1
        imin = i
        For j = i + 1 To N
            If pos(j) < pos(imin) Then imin = j
        Next
        If imin <> i Then
            tp = pos(i): pos(i) = pos(imin): pos(imin) = tp
            ts = pStart(i): pStart(i) = pStart(imin): pStart(imin) = ts
        End If
    Next
End Sub


'��һ����ʱ������ʼ
Private Function t0() As Double
    t0 = Timer
End Function

'��������ʱ�����׶κ�ʱ����д����ȴ��壩
Private Sub PF_Tick(ByVal pf As progressForm, ByRef t As Double, ByVal phase As String, Optional ByVal px As Long = -1)
    On Error Resume Next
    Dim dt As Double: dt = Timer - t: t = Timer
    If px < 0 Then
        pf.UpdateProgressBar pf.FrameProgress.width, "? " & phase & " ��ʱ " & Format(dt, "0.00") & " s"
    Else
        pf.UpdateProgressBar px, "? " & phase & " ��ʱ " & Format(dt, "0.00") & " s"
    End If
End Sub

'���������������ı������λ�ã�ֻ׷��һ����Ϣ
Private Sub PF_Log(ByVal pf As progressForm, ByVal msg As String)
    On Error Resume Next
    pf.UpdateProgressBar pf.FrameProgress.width, msg
End Sub


' ֻ�ڡ���ʱ>��ֵ���Ұ�����Ƶ��ʱдһ����������ǿ��ÿ�����������ˢ����
Private Sub PF_StepWarn(ByVal pf As progressForm, ByRef t As Double, _
                        ByVal tag As String, ByVal i As Long, ByVal N As Long, _
                        Optional ByVal warn As Double = 0.15, _
                        Optional ByVal sampleEvery As Long = 10, _
                        Optional ByVal basePx As Long = 70, _
                        Optional ByVal spanPx As Long = 80)
    On Error Resume Next
    Dim dt As Double: dt = Timer - t: t = Timer
    If (dt >= warn) And ((i <= 5) Or (i Mod sampleEvery = 0) Or (i = N)) Then
        Dim px As Long
        ' �����׶ν���ӳ�䵽 basePx~(basePx+spanPx)������������ν�������һ�£�
        px = basePx + CLng(spanPx * i / IIf(N = 0, 1, N))
        pf.UpdateProgressBar px, "�� �� " & i & "/" & N & " �� " & tag & " ��ʱ " & Format(dt, "0.000") & " s"
        DoEvents
    End If
End Sub

