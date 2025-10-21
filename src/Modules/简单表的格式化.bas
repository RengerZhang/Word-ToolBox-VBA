Attribute VB_Name = "�򵥱�ĸ�ʽ��"
Sub FormatTable()
    ' ��������
    Dim doc As Document
    Dim tbl As Table
    Dim titleRow As row
    Dim i As Integer, j As Integer ' ѭ���к���
    Dim cell As cell ' ������Ԫ��
    Dim minRowHeight As Single ' �и���Сֵ��0.6cm��
    
    ' 1. ����Ƿ�ѡ�б��
    On Error Resume Next
    Set tbl = Selection.Tables(1)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "����ѡ��һ����������д˺꣡", vbExclamation, "��ʾ"
        Exit Sub
    End If
    Set doc = ActiveDocument
    
    ' ����0.6���׶�Ӧ�İ�ֵ��1���ס�28.35����
    minRowHeight = CentimetersToPoints(0.6)
    
    ' 2. �����������и�ʽ��������ɫ��
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            Set cell = tbl.cell(i, j)
            With cell.Range
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
        Next j
    Next i
    
    ' 3. �����������
    tbl.AutoFitBehavior wdAutoFitWindow ' ��Ӧ���ڿ��
    tbl.AllowPageBreaks = True ' �����ҳ
    
    ' 4. �������е�Ԫ���ʽ�������޸ģ��߾�;��У�
    For i = 1 To tbl.rows.Count
        For j = 1 To tbl.Columns.Count
            Set cell = tbl.cell(i, j)
            
            ' ���õ�Ԫ��߾ࣨ���±߾�Ϊ0��
            With cell
                .TopPadding = 0 ' �ϱ߾�0��
                .BottomPadding = 0 ' �±߾�0��
                ' ���ұ߾ౣ��Ĭ�ϣ������������ӣ�.LeftPadding = 0 �� .RightPadding = 0��
            End With
            
            ' �������ָ�ʽ
            With cell.Range.Font
                .NameFarEast = "����" ' ��������
                .NameAscii = "Times New Roman" ' ����Times New Roman
                .Size = 10 ' ͳһ�ֺ�
            End With
            
            ' ����ˮƽ���У����Ҿ��У�
            cell.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
            
            ' ���ִ�ֱ���У��߶Ⱦ��У�
            cell.VerticalAlignment = wdCellAlignVerticalCenter
        Next j
    Next i
    
    ' 5. ���ñ����У���һ�У�
    Set titleRow = tbl.rows(1)
    titleRow.HeadingFormat = True ' ��ҳ�ظ�������
    titleRow.Range.Font.bold = True ' �����мӴ�
    
    ' 6. �����и߹��򣨺����޸ģ���Сֵ0.6cm��
    For i = 1 To tbl.rows.Count
        With tbl.rows(i)
            .HeightRule = wdRowHeightAtLeast ' �и�����Ϊָ��ֵ�����ݶ�ʱ�Զ���չ��
            .Height = minRowHeight ' ��Сֵ0.6cm
        End With
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
           "4. �����мӴֲ���ҳ�ظ�", vbInformation, "���"
End Sub

