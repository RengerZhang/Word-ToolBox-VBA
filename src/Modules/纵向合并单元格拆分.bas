Attribute VB_Name = "����ϲ���Ԫ����"
Sub Test���ѡ�еı��ĺϲ���Ԫ��()
    Dim tbl As Table
    Dim i, j, errorCount As Integer
    Dim ��ʼ������к� As Integer
    Dim zonghang As Integer, zonglie As Integer
    Dim currentCell As cell
    Dim mergedRows As Integer
    
    ' ��ȡѡ�еı��
    If Selection.Tables.Count = 0 Then
        MsgBox "����ѡ��һ�����", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    Set tbl = Selection.Tables(1) ' ��ȡѡ�еĵ�һ�����
    
    zonghang = tbl.rows.Count ' ������
    zonglie = tbl.Columns.Count ' ������
    
    
    ' ����ÿһ��
    For j = 1 To zonglie
        On Error Resume Next ' ���ô�����
        
        ' ����ÿһ��
        For i = 1 To zonghang
            Set currentCell = tbl.cell(i, j)
            
            ' �ж��Ƿ�Ϊ�ϲ���Ԫ��
            currentCell.Select
            If Err.Number <> 0 Then
                ' ��������ϲ���Ԫ�񣬼�¼�ϲ���Ԫ�����ʼλ��
                If errorCount = 0 Then
                    ��ʼ������к� = i - 1
                End If
                
                ' ���Ӵ������
                errorCount = errorCount + 1
                
                ' ���ϲ���Ԫ������ɫ
                currentCell.Shading.BackgroundPatternColor = wdColorYellow ' ����Ϊ��ɫ
                
                ' �������
                Err.Clear
            Else
                ' �����ǰ�ǷǺϲ���Ԫ�񣬴���ϲ���Ԫ��Ľ���
                If errorCount > 0 Then
                    ' ����ϲ���Ԫ���λ�úͺϲ�������
                    �ϲ����� = errorCount + 1
                    Debug.Print "�ϲ��ĵ�Ԫ���ڣ���" & ��ʼ������к� & "�У���" & j & "�У��ϲ���" & �ϲ����� & "��"
                    
                    ' ��¼�ϲ���Ԫ�����������
                    mergedRows = �ϲ�����
                    
                    ' ��ֺϲ���Ԫ��
                    Call ��ֵ�Ԫ��(CInt(��ʼ������к�), CInt(mergedRows), CInt(j)) ' ǿ��ת��Ϊ��������
                    
                    ' ���ô��������
                    errorCount = 0
                    ��ʼ������к� = 0
                End If
            End If
        Next i
        
        On Error GoTo 0 ' ���ô�����
    Next j
    
     ' ѭ��������ѡ���һ�в��������ּӴ�
    tbl.rows(1).Range.Font.bold = True
    Debug.Print "��һ�������ѼӴ�"
    
End Sub


' ��ֺϲ��ĵ�Ԫ��
Sub ��ֵ�Ԫ��(startRow As Integer, mergedRows As Integer, col As Integer)
    Dim currentCell As cell
    
    ' ѡ�е�ǰ�ϲ��ĵ�Ԫ��
    Set currentCell = Selection.Tables(1).cell(startRow, col)
    
    ' �����ֵĵ�Ԫ��λ��
    Debug.Print "���ڲ�֣���" & startRow & "�е�" & col & "�еĵ�Ԫ��"
    
    ' ѡ�кϲ���Ԫ��
    currentCell.Select
    
    ' ִ�в�ֲ�������ֺϲ���Ԫ��
    Selection.Cells.Split NumRows:=mergedRows, NumColumns:=1, MergeBeforeSplit:=False
    
    Debug.Print "�ϲ���Ԫ���Ѳ�֣���" & startRow & "�е�" & col & "�еĵ�Ԫ��"
End Sub

