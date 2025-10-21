Attribute VB_Name = "��ʽ_��׼�������ʽ"
Option Explicit

'========================================
' Ԥ�����������ֵ�Ԫ�� = �ܵ�Ԫ�� ? ��ͼ��Ԫ�񡱹�����
' �ж���
'   n = ͼƬ������Inline + Shape��
'   �� �������ֵ�Ԫ�� �� (n + 1) �� ��ͼƬ��λ�������� �� ����׼�����ʽ��
'   ��ͼ��n=0���� ֱ�ӡ���׼�����ʽ��
' ������ڡ��������ڡ�(Ctrl+G)���ÿ����ϸ
'========================================
Public Sub Ԥ����_���ͼƬ������ͨ��()
    '��һ����ʽ��
    Const S_TABLE_PIC As String = "ͼƬ��λ��"
    Const S_TABLE_NOR As String = "��׼�����ʽ"

    '��������֤��ʽ���ڣ�������ۣ�
    Dim doc As Document: Set doc = ActiveDocument
    EnsureTableStyleOnly doc, S_TABLE_PIC
    EnsureTableStyleOnly doc, S_TABLE_NOR

    '����������ж�
    Dim i As Long, tb As Table
    Dim nInline As Long, nShape As Long, nImgObj As Long
    Dim totalCells As Long, imgCellCnt As Long, txtCellEst As Long
    Dim threshold As Long, applied As String, imgCoords As String

    For i = 1 To doc.Tables.Count
        Set tb = doc.Tables(i)

        ' 1) ͼƬ����������n��
        nInline = tb.Range.InlineShapes.Count
        nShape = SafeShapeCount_InRange(tb.Range)
        nImgObj = nInline + nShape

        ' 2) ��Ԫ���ͳ�ƣ��ܵ�Ԫ���� & ��ͼ��Ԫ����
        totalCells = tb.Range.Cells.Count                           ' ���ݺϲ���Ԫ��
        imgCellCnt = CountImageCells(tb, imgCoords)                 ' ��Ԫ�����Ƿ�ͼ
        txtCellEst = totalCells - imgCellCnt                        ' �������ֵ�Ԫ����

        ' 3) �ж�����ֵ = n + 1
        threshold = nImgObj + 1
        If (nImgObj > 0) And (txtCellEst <= threshold) Then
            tb.Style = S_TABLE_PIC
            applied = S_TABLE_PIC
        Else
            tb.Style = S_TABLE_NOR
            applied = S_TABLE_NOR
        End If

        ' 4) �������
        Debug.Print "��#" & i & _
                    " | �ߴ�=" & tb.rows.Count & "x" & tb.Columns.Count & _
                    " | �ܵ�Ԫ��=" & totalCells & _
                    " | ͼƬ���� n=Inline:" & nInline & "+Shape:" & nShape & "=" & nImgObj & _
                    " | ��ͼ��Ԫ��=" & imgCellCnt & _
                    " | �������ֵ�Ԫ��=" & txtCellEst & " ��? ��ֵ(n+1)=" & threshold & _
                    " | �ж�=" & applied & _
                    IIf(Len(imgCoords) > 0, " | ��ͼ����:" & imgCoords, "")
    Next i

    MsgBox "Ԥ������ɣ������ܵ�Ԫ��?��ͼ��Ԫ�񡱼������ֵ�Ԫ�񣩡�������������ڡ�", vbInformation
End Sub

'========================================
' ���ߣ�ͳ�� Range �ڡ�������״������������״ʱ������
'========================================
Private Function SafeShapeCount_InRange(ByVal rng As Range) As Long
    On Error Resume Next
    SafeShapeCount_InRange = rng.ShapeRange.Count
    On Error GoTo 0
End Function

'========================================
' ���ߣ�ͳ�ơ���ͼ��Ԫ�������������������嵥 "(r,c),(r,c)..."
' ���򣺵�Ԫ���� InlineShapes.Count + ShapeRange.Count > 0 ����Ϊ����ͼ��
'========================================
Private Function CountImageCells(ByVal tb As Table, ByRef coords As String) As Long
    Dim c As cell, n As Long, buf As String
    For Each c In tb.Range.Cells
        If (c.Range.InlineShapes.Count > 0) Or (SafeShapeCount_InRange(c.Range) > 0) Then
            n = n + 1
            If Len(buf) > 0 Then buf = buf & ","
            ' Word �� Cell ����֧�� RowIndex / ColumnIndex
            buf = buf & "(" & c.rowIndex & "," & c.ColumnIndex & ")"
        End If
    Next c
    coords = buf
    CountImageCells = n
End Function

'========================================
' ���ף���֤��ʽ������Ϊ�������ʽ����������ۣ�
'========================================
Private Sub EnsureTableStyleOnly(ByVal doc As Document, ByVal styleName As String)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If Not st Is Nothing Then
        If st.Type <> wdStyleTypeTable Then
            st.Delete
            Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
        End If
    Else
        Set st = doc.Styles.Add(name:=styleName, Type:=wdStyleTypeTable)
    End If
End Sub


