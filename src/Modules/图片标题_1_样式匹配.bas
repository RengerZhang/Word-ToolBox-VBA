Attribute VB_Name = "ͼƬ����_1_��ʽƥ��"
Option Explicit

'========================
' �Լ죺�Ƿ��ѳɹ��������֡������ʽ��
' ˵������������ĵ���ʽ�������ĵ���ͷ��ʱ����1��1����������֤�����ɾ��
'========================
Public Sub �Լ�_�����ʽ����״̬()
    '��һ����������ʽ���������ڿɼ�������ģ�鼶��ͻ��
    Const S_NORMAL As String = "��׼�����ʽ"
    Const S_PIC    As String = "ͼƬ��λ��"

    '������׼������
    Dim doc As Document: Set doc = ActiveDocument
    Dim okNormal As Boolean, okPic As Boolean
    Dim msg As String: msg = "�����ʽ����״̬��" & vbCrLf

    '������������ + ���ͣ�����Ϊ�������ʽ�������
    okNormal = StyleExistsAsTable(doc, S_NORMAL)
    okPic = StyleExistsAsTable(doc, S_PIC)

    msg = msg & " - [" & S_NORMAL & "] " & IIf(okNormal, "�Ѵ��ڣ������ʽ��", "δ�ҵ�") & vbCrLf
    msg = msg & " - [" & S_PIC & "] " & IIf(okPic, "�Ѵ��ڣ������ʽ��", "δ�ҵ�")

    '���ģ��ԡ�ͼƬ��λ����һ�Ρ�ʵ�⡱���½���ʱ�������ʽ����ȡ�߿�/�ڱ߾�
    If okPic Then
        Dim rng As Range, tb As Table
        Dim passBorders As Boolean, passPadding As Boolean

        Set rng = doc.Range(0, 0)
        rng.Collapse wdCollapseStart

        On Error Resume Next
        doc.UndoRecord.StartCustomRecord "ͼƬ��λ��-�Լ�"
        On Error GoTo 0

        Set tb = doc.Tables.Add(rng, 1, 1)
        tb.Style = S_PIC

        passBorders = (tb.Borders.enable = False _
                    And tb.Borders(wdBorderTop).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderBottom).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderLeft).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderRight).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone _
                    And tb.Borders(wdBorderVertical).LineStyle = wdLineStyleNone)

        passPadding = (tb.TopPadding = 0 And tb.BottomPadding = 0 _
                    And tb.LeftPadding = 0 And tb.RightPadding = 0)

        ' ɾ����ʱ��
        tb.Range.Delete

        On Error Resume Next
        doc.UndoRecord.EndCustomRecord
        On Error GoTo 0

        msg = msg & vbCrLf & vbCrLf & "ͼƬ��λ��ʵ�����ԣ���" _
            & vbCrLf & " - ����ȫ�أ� " & BoolCN(passBorders) _
            & vbCrLf & " - �ڱ߾�Ϊ0�� " & BoolCN(passPadding)
    End If

    '���壩����
    MsgBox msg, IIf(okNormal And okPic, vbInformation, vbExclamation), "��ʽ�����Լ�"
End Sub

'========================
' ��ʾ���ѡ���ѡ��񡱱��ΪͼƬ��Ŀ����֤��ֱ�ۣ�
'========================
Public Sub һ������ѡ�����ΪͼƬ��()
    Const S_PIC As String = "ͼƬ��λ��"
    If Not StyleExistsAsTable(ActiveDocument, S_PIC) Then
        MsgBox "δ�ҵ���ʽ��" & S_PIC & "��������ִ����ġ�һ��������ʽ����", vbExclamation
        Exit Sub
    End If

    If Selection.Information(wdWithInTable) Then
        Selection.Tables(1).Style = S_PIC
        MsgBox "�ѽ���ѡ�������Ϊ��ͼƬ��λ������Ŀ�⣺�ޱ߿��ڱ߾�Ϊ0��", vbInformation
    Else
        MsgBox "���Ȱѹ��ŵ�Ҫ���Եı���", vbExclamation
    End If
End Sub

'========================
' ���ߣ��ж�ĳ��ʽ�Ƿ������Ϊ�������ʽ��
'========================
Private Function StyleExistsAsTable(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    StyleExistsAsTable = (Not st Is Nothing And st.Type = wdStyleTypeTable)
End Function

'========================
' ���ߣ�����ֵ����
'========================
Private Function BoolCN(ByVal flag As Boolean) As String
    BoolCN = IIf(flag, "��", "��")
End Function


