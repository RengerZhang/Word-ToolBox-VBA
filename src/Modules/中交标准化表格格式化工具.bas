Attribute VB_Name = "�н���׼������ʽ������"
Sub AdjustTableFormat()
    Dim oRow As row
    Dim oCell As cell
    Dim tb As Table
    Dim myParagraph As Paragraph, N As Integer
    
    ' ���á���׼�������ʽ������
    EnsureStandardTableStyle
    
    Selection.Tables(1).Select
    Selection.Style = "��׼�������ʽ"
    
    
    tbl = Selection.Tables(1)
    Set tb = Selection.Tables(1)
    
    ' ����Ƿ�ѡ���˱��
    On Error Resume Next
    Set tbls = Selection.Tables  ' ��ȡѡ�������еı�񼯺�
    On Error GoTo 0
    
    ' ���û��ѡ�б����ʾ�û����˳�
    If tbls.Count = 0 Then
        MsgBox "����ѡ��һ����������б��꣡", vbExclamation, "��ѡ�б��"
        Exit Sub
    End If
    
    
    '    Selection.ParagraphFormat.Reset
    '    tb.AutoFitBehavior (wdAutoFitContent)
    tb.AutoFitBehavior (wdAutoFitWindow)
    tb.TopPadding = PixelsToPoints(0, True) '�����ϱ߾�Ϊ0
    tb.BottomPadding = PixelsToPoints(0, True) '�����±߾�Ϊ0
    tb.LeftPadding = PixelsToPoints(0, True) '�����ϱ߾�Ϊ0
    tb.RightPadding = PixelsToPoints(0, True) '�����±߾�Ϊ0
    
    '��ʽ�����
    tbl.Select
    With Selection
        .rows.alignment = wdAlignRowCenter
        .rows.WrapAroundText = False
        .Font.NameFarEast = ""
        .Font.NameAscii = ""
        .Range.bold = False
        .Range.Font.NameAscii = "Times New Roman"
        .Range.Font.NameFarEast = "����"
        .Range.Font.Size = 10.5
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter ' ��Ԫ��ֱ����
        .ParagraphFormat.CharacterUnitFirstLineIndent = 0
        .ParagraphFormat.alignment = wdAlignParagraphCenter ' ���ж���
        .ParagraphFormat.SpaceBefore = 0 ' ��ǰ
        .ParagraphFormat.SpaceAfter = 0 ' �κ�
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle ' �����о�
        .ParagraphFormat.CharacterUnitFirstLineIndent = 0 ' �������
        .ParagraphFormat.LeftIndent = 0 ' �������
        .ParagraphFormat.RightIndent = 0 ' �Ҳ�����
        .Shading.BackgroundPatternColor = wdColorAutomatic ' �����ɫ
        .rows.HeightRule = wdRowHeightAtLeast
        .rows.Height = CentimetersToPoints(0.6)
    End With
    
    '  ������Ԫ�������ڿ���,Ϊ�˷�ֹ�ϲ���Ԫ������
    For Each oCell In tbl.Cells
        oCell.Select
        With Selection
            .Borders.OutsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineWidth = wdLineWidth050pt
        End With
        
        Selection.SelectRow
        Selection.rows.AllowBreakAcrossPages = enable
        
        N = 1
        For Each myParagraph In Selection.Paragraphs
            If Len(Trim(myParagraph.Range)) = 1 Then
                myParagraph.Range.Delete
                N = N + 1
            End If
        Next
        
    Next oCell
    
    
    '  ѡ�б�����������
    With tbl.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideLineWidth = wdLineWidth150pt ' ���1.5��
        .OutsideColor = wdColorBlack
    End With
    
    '  �������мӴֲ���ÿҳ�ظ�
    tbl.Select
    Selection.rows.HeadingFormat = False
    
    tbl.Cells(1).Select
    Selection.SelectRow
    Selection.Range.bold = True
    Selection.rows.HeadingFormat = True
    
    ' ��ϸ������ʾ��Ϣ
    MsgBox "����ʽ������ɣ���ִ�����²�����" & vbCrLf & _
           "1. Ӧ��""���""��ʽ" & vbCrLf & _
           "2. �Զ���Ӧ���ڿ��" & vbCrLf & _
           "3. ���õ�Ԫ�����±߾�Ϊ0" & vbCrLf & _
           "4. �о��ж��룬ȡ�����ֻ���" & vbCrLf & _
           "5. �������ã�����(����)��Times New Roman(Ӣ��)��10.5����" & vbCrLf & _
           "6. ��Ԫ�����ݴ�ֱ��ˮƽ����" & vbCrLf & _
           "7. �������ã������о࣬��ǰ���࣬������" & vbCrLf & _
           "8. �����ɫ���и���Ϊ��Сֵ0.6cm" & vbCrLf & _
           "9. �����ڿ���(0.5��)�������(1.5��)" & vbCrLf & _
           "10. ��ֹ����п�ҳ����" & vbCrLf & _
           "11. ɾ����Ԫ���ڿն���" & vbCrLf & _
           "12. ���мӴֲ�����Ϊÿҳ�ظ�������", _
           vbInformation, "�������"
    
End Sub
