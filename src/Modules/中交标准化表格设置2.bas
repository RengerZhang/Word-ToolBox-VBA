Attribute VB_Name = "�н���׼���������2"
Sub FormatTable()
    ' ��������
    Dim tbl As Table
    Dim titleRow As row
    Dim i As Integer, j As Integer, k As Integer
    Dim minRowHeight As Single
    Dim currentCell As cell
    Dim targetRow As row
    Dim processedRows As New Collection ' ���ڴ洢�Ѵ�����У������ظ�����
    Dim originalRange As Range ' ���ڱ���ԭʼѡ��Χ
    Dim tableRange As Range ' ������巶Χ
    Dim firstRowRange As Range ' ��һ�з�Χ
    
    ' ����ԭʼѡ��Χ�����⴦������иı��û�ѡ��
    Set originalRange = Selection.Range
    
    ' 1. ����Ƿ�ѡ�б��
    On Error Resume Next
    Set tbl = Selection.Tables(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "����ѡ��һ����������д˺꣡", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    ' ����0.6���׶�Ӧ�İ�ֵ��1���ס�28.35����
    minRowHeight = CentimetersToPoints(0.6)
    
    ' 2. �����������и�ʽ��������ɫ��
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            ' ����ϲ���Ԫ����ܵ��µĵ�Ԫ�񲻴�������
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                With currentCell.Range
                    ' �����������
                    With .Font
                        .NameFarEast = ""
                        .NameAscii = ""
                        .bold = False
                        .Italic = False
                        .Underline = wdUnderlineNone
                        .Color = wdColorAutomatic
                        .Size = 10
                    End With
                    
                    ' �����������
                    With .ParagraphFormat
                        .alignment = wdAlignParagraphLeft
                        .LeftIndent = 0
                        .RightIndent = 0
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .LineSpacingRule = wdLineSpaceSingle
                    End With
                    
                    ' �����Ԫ���ɫ
                    .Shading.BackgroundPatternColor = wdColorAutomatic
                End With
                
                ' �ͷŶ���
                Set currentCell = Nothing
            End If
        Next j
    Next i
    
    ' 3. �����������
    tbl.AutoFitBehavior wdAutoFitWindow ' ��Ӧ���ڿ��
    tbl.AllowPageBreaks = True ' �����ҳ
    
    ' 4. �������е�Ԫ���ʽ���߾ࡢ���ָ�ʽ�����У�
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                ' ���õ�Ԫ��߾ࣨ���±߾�Ϊ0��
                With currentCell
                    .TopPadding = 0 ' �ϱ߾�0��
                    .BottomPadding = 0 ' �±߾�0��
                End With
                
                ' �������ָ�ʽ
                With currentCell.Range.Font
                    .NameFarEast = "����" ' ��������
                    .NameAscii = "Times New Roman" ' ����Times New Roman
                    .Size = 10 ' ͳһ�ֺ�
                End With
                
                ' ����ˮƽ����
                currentCell.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
                
                ' ���ִ�ֱ����
                currentCell.VerticalAlignment = wdCellAlignVerticalCenter
                
                ' �ͷŶ���
                Set currentCell = Nothing
            End If
        Next j
    Next i
    
    ' 2. ͨ�����Χ��ȡ��һ�У�����ֱ��ʹ�� Rows(1)
    On Error Resume Next
    ' ��ȡ������巶Χ
    Set tableRange = tbl.Range
    ' �ӱ��Χ�н�ȡ��һ�еķ�Χ���ؼ��Ľ���
    Set firstRowRange = tableRange.rows(1).Range
    On Error GoTo 0
    
    If Not firstRowRange Is Nothing Then
        ' 3. �ӵ�һ�з�Χ����ȡ�ж���
        On Error Resume Next
        Set titleRow = firstRowRange.rows(1)
        On Error GoTo 0
'
        ' 4. ��֤�ж������ñ���������
        If Not titleRow Is Nothing Then
            titleRow.HeadingFormat = True ' ��ҳ�ظ�
            titleRow.Range.Font.bold = True ' �Ӵ�
        Else
            ' ����������ֱ�Ӳ�����һ�з�Χ�ĸ�ʽ
            firstRowRange.Font.bold = True
            ' ��ҳ�ظ�������Ҫ�ж��󣬴˴���ʾ����ʧЧ
            Debug.Print "���棺�޷����ÿ�ҳ�ظ�������ɼӴ�"
        End If
    Else
        ' �ռ�������ֱ��ͨ����Ԫ��Χ���ø�ʽ
        On Error Resume Next
        ' ��������һ�е�һ����Ԫ�����ڵ��з�Χ
        tbl.cell(1, 1).Range.rows(1).Font.bold = True
        tbl.cell(1, 1).Range.rows(1).HeadingFormat = True
        On Error GoTo 0
        Debug.Print "ʹ�õ�Ԫ��Χ�������ñ�����"
    End If
    

    
    ' �ָ�ԭʼѡ��Χ
    originalRange.Select
    
    ' 6. �����и߹��򣨴���ϲ���Ԫ����з������⣩
    ' �ռ�����Ψһ�У������ظ�����
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            On Error Resume Next
            Set currentCell = tbl.cell(i, j)
            On Error GoTo 0
            
            If Not currentCell Is Nothing Then
                ' �ɿ����л�ȡ��ʽ
                Dim rowIndex As Integer
                rowIndex = currentCell.rowIndex
                On Error Resume Next
                Set targetRow = tbl.rows(rowIndex)
                On Error GoTo 0
                
                If Not targetRow Is Nothing Then
                    ' ������Ƿ��Ѵ���
                    Dim isExists As Boolean
                    isExists = False
                    For Each existingRow In processedRows
                        If existingRow = targetRow.Index Then
                            isExists = True
                            Exit For
                        End If
                    Next
                    If Not isExists Then
                        On Error Resume Next
                        processedRows.Add targetRow.Index, CStr(targetRow.Index)
                        ' �����и�
                        targetRow.HeightRule = wdRowHeightAtLeast
                        targetRow.Height = minRowHeight
                        On Error GoTo 0
                    End If
                Else
                    ' ��¼������Ϣ
                    Debug.Print "��ȡ�ж���ʧ�ܣ���������" & rowIndex & "����Ԫ��λ�ã���" & i & "�е�" & j & "��"
                End If
                
                ' �ͷŶ���
                Set currentCell = Nothing
                Set targetRow = Nothing
            End If
        Next j
    Next i
    
    ' 7. ���ñ߿����1.5�����ڿ�0.5����
    With tbl.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideLineWidth = wdLineWidth150pt ' 1.5��
        .OutsideColor = wdColorBlack
        
        .InsideLineStyle = wdLineStyleSingle
        .InsideLineWidth = wdLineWidth050pt ' 0.5��
        .InsideColor = wdColorBlack
    End With
    
    ' �����ʾ
    MsgBox "����ʽ������ɣ�" & vbCrLf & _
           "1. �и���Сֵ0.6cm�����ݶ�ʱ�Զ���չ��" & vbCrLf & _
           "2. ��Ԫ�����±߾�Ϊ0������ˮƽ+��ֱ����" & vbCrLf & _
           "3. ���1.5�����ڿ�0.5�����������壬����Times New Roman" & vbCrLf & _
           "4. �����У����ϲ��У��Ӵֲ���ҳ�ظ�", vbInformation, "���"
End Sub


