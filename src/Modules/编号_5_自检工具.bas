Attribute VB_Name = "���_5_�Լ칤��"
Option Explicit

'==========================================================
' �����Լ�С���ߣ������棩
' �仯�㣺
'   1) �����ĵ���A4 ����ҳ�߾� 2cm
'   2) ȫ��Ĭ�����壺����=���壬����/����=Times New Roman���ֺ� 10.5
'   3) ������ġ�����Ų�����д������ Unicode ���ɣ�֧�ֶ�����ϣ�
'   4) ���ӡ�ƥ�����м������У����ٶ�λ 5~7 ��û���е�ԭ��
'
' ������
'   - Public Function ��ȡ���м������()  ' ���Բ��裨�ۣ�
'   - ��ʹ���� Mod�������ģ��뱣�֡�������򡱵��߼�һ�£�����ͬ���滻��
'==========================================================
Sub �����Լ�_���ɱ���()
    Dim srcDoc As Document   ' �� Ҫ����Դ�ĵ�
    Dim rptDoc As Document   ' �� �½����Լ챨��
    Dim cfg As Variant
    Dim i As Long, N As Long
    Dim rng As Range
    Dim tbl As Table
    Dim �� As Long
    Dim ��ʽ�� As String, ��Ÿ�ʽ As String
    Dim �����ʽֵ As Long, ����cm As Single
    Dim ƥ������ As String, ɾ������ As String
    Dim ��ʽ�Ƿ���� As String
    Dim �ֹ����� As Long, �Զ����� As Long
    Dim ��ʽ�嵥 As Variant, s As Variant
    Dim ɾ��ȫ�� As Variant, p As Variant
    Dim ����ӳ�� As Variant, t As Long

    ' 1) ���½�����֮ǰ��������ץס��Դ�ĵ���
    Set srcDoc = ActiveDocument

    ' 2) ��ȡ�������ã����Բ�������
    cfg = ��ȡ���м������()
    If IsEmpty(cfg) Then
        MsgBox "δ��ȡ��������ã���ȡ���м������() ����Ϊ�գ���", vbExclamation
        Exit Sub
    End If
    N = UBound(cfg, 1)

    ' 3) �½������ĵ������ı� ActiveDocument�����Ժ���һ���� srcDoc/rptDoc��
    Set rptDoc = Documents.Add
    With rptDoc.PageSetup
        .Orientation = wdOrientLandscape
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With
    With rptDoc.Styles(wdStyleNormal).Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 10.5
    End With
    rptDoc.content.Style = rptDoc.Styles(wdStyleNormal)

    ' 4) ����
    Set rng = rptDoc.Range(0, 0)
    rng.text = "�����Լ챨��" & vbCrLf & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf

    ' 5) ������ӡ��Զ���ż������У������������ࣩ
    Set tbl = rptDoc.Tables.Add(Range:=rng.Duplicate, NumRows:=N + 1, NumColumns:=10)
    With tbl
        .AllowAutoFit = True
        .AutoFitBehavior wdAutoFitWindow
        .rows(1).Range.bold = True
        .rows(1).Shading.BackgroundPatternColor = wdColorGray20
        .Borders.enable = True
        .Range.Font.NameFarEast = "����"
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameOther = "Times New Roman"
        .Range.Font.Size = 10.5

        .cell(1, 1).Range.text = "����"
        .cell(1, 2).Range.text = "��ʽ��"
        .cell(1, 3).Range.text = "��ʽ�Ƿ���ڣ�Դ�ĵ���"
        .cell(1, 4).Range.text = "��Ÿ�ʽ"
        .cell(1, 5).Range.text = "�����ʽ"
        .cell(1, 6).Range.text = "����(cm)"
        .cell(1, 7).Range.text = "����ƥ�����򣨶��ף�"
        .cell(1, 8).Range.text = "ɾ��������򣨶��ף�"
        .cell(1, 9).Range.text = "���м���������/�ֹ���"
        .cell(1, 10).Range.text = "���м������Զ����=�ü���"
        
        
        '�����п����Ԥ�裨�ܿ� = ҳ���� - ���ұ߾ࣩ
        Dim cw As Single
        Dim r(1 To 10) As Double
        Dim k As Long
        
        ' �ر�����Ӧ�����̶��������
        .AllowAutoFit = False
        
        ' ҳ����ÿ�ȣ���λ��pt��
        cw = rptDoc.PageSetup.PageWidth - rptDoc.PageSetup.LeftMargin - rptDoc.PageSetup.RightMargin
        
        ' ������1..10 ������Ϊ
        ' ����, ��ʽ��, �Ƿ����, ��Ÿ�ʽ, �����ʽ, ����, ����ƥ������, ɾ���������, ����ǰ׺ƥ����, ��ټ���ƥ����
        r(1) = 0.06: r(2) = 0.12: r(3) = 0.1: r(4) = 0.12: r(5) = 0.1
        r(6) = 0.08: r(7) = 0.18: r(8) = 0.18: r(9) = 0.03: r(10) = 0.03
        
        ' Ӧ���п����������ã�
        For k = 1 To 10
            .Columns(k).width = cw * r(k)
        Next k

    End With

    ' 6) �����ȫ���� srcDoc Ϊ׼��
    For i = 1 To N
        ��ʽ�� = CStr(cfg(i, 1))
        ��Ÿ�ʽ = CStr(cfg(i, 2))
        �����ʽֵ = CLng(cfg(i, 3))
        ����cm = CSng(cfg(i, 4))

        ��ʽ�Ƿ���� = IIf(��ʽ����(srcDoc, ��ʽ��), "��", "��")

        ƥ������ = �������ƥ�����_�Լ�v2(�����ʽֵ, ��Ÿ�ʽ)
        ɾ������ = ����ɾ������_�Լ�v2(�����ʽֵ, ��Ÿ�ʽ)

        �ֹ����� = ͳ���ĵ�������(srcDoc, ƥ������)
        �Զ����� = ͳ���Զ���ż���(srcDoc, i)

        tbl.cell(i + 1, 1).Range.text = CStr(i)
        tbl.cell(i + 1, 2).Range.text = ��ʽ��
        tbl.cell(i + 1, 3).Range.text = ��ʽ�Ƿ����
        tbl.cell(i + 1, 4).Range.text = ��Ÿ�ʽ
        tbl.cell(i + 1, 5).Range.text = ӳ������ʽ��(�����ʽֵ)
        tbl.cell(i + 1, 6).Range.text = Format(����cm, "0.##")
        tbl.cell(i + 1, 7).Range.text = ƥ������
        tbl.cell(i + 1, 8).Range.text = ɾ������
        tbl.cell(i + 1, 9).Range.text = CStr(�ֹ�����)
        tbl.cell(i + 1, 10).Range.text = CStr(�Զ�����)
    Next i

    ' 7) ׷���嵥��һ���� srcDoc Ϊ׼
    Set rng = rptDoc.Range(rptDoc.content.End - 1, rptDoc.content.End - 1)
    
    '�����ڱ���׷�ӡ�ע�͡�˵������ָ��
    Dim noteRng As Range
    Set noteRng = rptDoc.Range(tbl.Range.End, tbl.Range.End)
    noteRng.InsertParagraphAfter
    noteRng.Collapse wdCollapseEnd
    noteRng.text = _
        "ע�ͣ�" & vbCrLf & _
        "? ����ǰ׺ƥ�����������򣩣�ͳ��Դ�ĵ��У��ü���Ӧ�ġ����ױ����̬�����ڶ����ı���ͷ������ƥ�䵽��������" & vbCrLf & _
        "  �����ı�ǰ׺���Զ���ŵ����ֲ��ڶ����ı����˲��������С�" & vbCrLf & _
        "? ��ټ���ƥ��������ListLevel����ͳ��Դ�ĵ��У�ʹ�á��༶��ٱ�š��Ҽ�����ڸü��Ķ���������" & vbCrLf & _
        "  ������ı��޹أ�ֱ�����ݶ���� ListLevelNumber �ж���" & vbCrLf & _
        "��ⷽʽ��ʾ������" & vbCrLf & _
        "  - ǰ׺ƥ������0������ƥ����>0���ü����������Զ���ţ���������" & vbCrLf & _
        "  - ǰ׺ƥ����>0������ƥ����=0���ü���Ϊ�ֹ���ţ�������ִ�С�����ƥ�䡱�롰�Զ��༶��š���" & vbCrLf & _
        "  - ���߶���ͬ�������ֹ�ǰ׺�����Զ���ţ�����ִ�С�ȥ���ֹ���š���" & vbCrLf & _
        "  - ���߶�С�������ʽ�Ƿ��Ѵ���/Ӧ�ã�������̬�Ƿ�������һ�¡�" & vbCrLf & vbCrLf

    
    rng.InsertAfter vbCrLf & "��A��Ŀ����ʽ����Դ�ĵ����ڣ�" & vbCrLf
    ��ʽ�嵥 = ��ȡ��ʽ������_����ĵ�(srcDoc, True)
    If IsArray(��ʽ�嵥) Then
        For Each s In ��ʽ�嵥
            rng.InsertAfter " - " & CStr(s) & vbCrLf
        Next
    Else
        rng.InsertAfter "(��)" & vbCrLf
    End If

    rng.InsertAfter vbCrLf & "��B��ɾ���ֹ���Ź��򼯣����ף�" & vbCrLf
    ɾ��ȫ�� = ����ɾ����Ź���()
    If IsArray(ɾ��ȫ��) Then
        For Each p In ɾ��ȫ��
            rng.InsertAfter " - " & CStr(p) & vbCrLf
        Next
    Else
        rng.InsertAfter "(��)" & vbCrLf
    End If

    rng.InsertAfter vbCrLf & "��C������ƥ����򼯣�pattern �� style��" & vbCrLf
    ����ӳ�� = ���ɱ���ƥ�����()
    If IsArray(����ӳ��) Then
        For t = LBound(����ӳ��, 1) To UBound(����ӳ��, 1)
            rng.InsertAfter " - " & CStr(����ӳ��(t, 1)) & "  ��  " & CStr(����ӳ��(t, 2)) & vbCrLf
        Next t
    Else
        rng.InsertAfter "(��)" & vbCrLf
    End If

    MsgBox "�����Լ챨�������ɣ���Դ�ĵ�Ϊ�ھ�����", vbInformation
End Sub

' ͳ�ơ��Զ���ż��� == targetLevel���Ķ�������
' ��ͳ�ƴ�ٱ�ţ�Outline Numbering����������Ŀ�������ȥ
Private Function ͳ���Զ���ż���(ByVal doc As Document, ByVal targetLevel As Long) As Long
    Dim p As Paragraph
    Dim c As Long, lvl As Long

    For Each p In doc.Paragraphs
        On Error Resume Next
        If p.Range.ListFormat.ListType = wdListOutlineNumbering Then   ' ֻ�϶༶��ٱ��
            lvl = p.Range.ListFormat.ListLevelNumber
            If lvl = targetLevel Then c = c + 1
        End If
        On Error GoTo 0
    Next

    ͳ���Զ���ż��� = c
End Function

'���� �Լ�棺ɾ���������������ı���һ�£�
Private Function ����ɾ������_�Լ�v2(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = ͳ��ռλ��(numFormat)
    Dim punct As String: punct = "[��,��:������.\-���C]"

    If numStyle = wdListNumberStyleNumberInCircle Then
        ����ɾ������_�Լ�v2 = "^[ \t]*[" & ��������ż�() & "]\s*"
        Exit Function
    End If
    If InStr(numFormat, "��%") > 0 Or InStr(numFormat, "(%") > 0 Then
        ����ɾ������_�Լ�v2 = "^[ \t]*[��(]\s*\d+\s*[)��]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If Right$(Trim$(numFormat), 1) = "��" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            ����ɾ������_�Լ�v2 = "^[ \t]*\d+\s*[)��]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If
    If c >= 2 Then
        ����ɾ������_�Լ�v2 = "^[ \t]*\d+(?:\s*[\.����]\s*\d+){1,}\s*(?:[\.����])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If c = 1 Then
        ����ɾ������_�Լ�v2 = "^[ \t]*\d+(?!\s*[)��])\s*(?:[\.����]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If
    ����ɾ������_�Լ�v2 = "^[ \t]*\d+[ ��\t]+"
End Function


'���� �Լ�棺����ƥ��������������ı���һ�£�
Private Function �������ƥ�����_�Լ�v2(ByVal numStyle As Long, ByVal numFormat As String) As String
    Dim c As Long: c = ͳ��ռλ��(numFormat)
    Dim punct As String: punct = "[��,��:������.\-���C]"

    If numStyle = wdListNumberStyleNumberInCircle Then
        �������ƥ�����_�Լ�v2 = "^[ \t]*[" & ��������ż�() & "]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If InStr(numFormat, "��%") > 0 Or InStr(numFormat, "(%") > 0 Then
        �������ƥ�����_�Լ�v2 = "^[ \t]*[��(][ \t]*\d+[ \t]*[)��]\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If Right$(Trim$(numFormat), 1) = "��" Or Right$(Trim$(numFormat), 1) = ")" Then
        If InStr(numFormat, "%") > 0 Then
            �������ƥ�����_�Լ�v2 = "^[ \t]*\d+[ \t]*[)��]\s*(?:" & punct & "\s*)?"
            Exit Function
        End If
    End If
    If c >= 2 Then
        �������ƥ�����_�Լ�v2 = "^[ \t]*\d+(?:\s*[\.����]\s*\d+){1,}\s*(?:[\.����])?\s*(?:" & punct & "\s*)?"
        Exit Function
    End If
    If c = 1 Then
        �������ƥ�����_�Լ�v2 = "^[ \t]*\d+(?!\s*[)��])\s*(?:[\.����]\s*)?(?:" & punct & "\s*)?"
        Exit Function
    End If
    �������ƥ�����_�Լ�v2 = "^[ \t]*\d+([ ��\t]+|$)"
End Function

