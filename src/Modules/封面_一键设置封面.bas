Attribute VB_Name = "����_һ�����÷���"
Option Explicit
Public Const SIGN_MIN_LINESPACE_PT As Single = 4        ' ȫ���̿ɼ�

'===============================
' �������ɣ�����������Ҫ��
'===============================
Public Sub ���ɷ���_��������()
    '������һ���İ�������ģ�
    Dim ��_��Ŀ�� As String
    ��_��Ŀ�� = "������������ MHPO-1403 ��Ԫ 73-04" & vbCrLf & "�ؿ�����(��Ǩ)����ס����Ŀ"

    Dim ��_������ As String
    ��_������ = "����֧������ˮ����������" & vbCrLf & "ר��ʩ������"

    ' �ϲ��������λ+��Ŀ����
    Dim ��_��� As String
    ��_��� = "�н������ֶ��ȹ������޹�˾" & vbCrLf & _
             "������������ MHPO-1403 ��Ԫ 73-04 �ؿ�����" & vbCrLf & "(��Ǩ)����ס����Ŀ����"

    Dim ��_���� As String: ��_���� = "2025��09��XX��"


    ' �����߶ȣ�����Ϊ��> ���С��İ�ȫֵ
    Const H_��Ŀ��_mm As Single = 30
    Const H_������_mm As Single = 35
    Const H_ǩ�ֿ�_mm As Single = 50
    Const H_���_mm  As Single = 35
    Const H_����_mm   As Single = 12
    
     '��������������/λ�ò�����������ԭ�߼���
    Const �ڱ߾�_mm As Single = 0
    
    Dim y_mm As Single: y_mm = 25
    
    Dim Y_��Ŀ��_mm As Single: Y_��Ŀ��_mm = y_mm
    y_mm = y_mm + H_��Ŀ��_mm + 12
    
    Dim Y_������_mm As Single: Y_������_mm = y_mm
    y_mm = y_mm + H_������_mm + 25
    
    Dim Y_ǩ�ֿ�_mm As Single: Y_ǩ�ֿ�_mm = y_mm
    y_mm = y_mm + H_ǩ�ֿ�_mm + 18
    
    Dim Y_���_mm As Single: Y_���_mm = y_mm
    y_mm = y_mm + H_���_mm
    
    Dim Y_����_mm As Single: Y_����_mm = y_mm


    ' ����ǩ�ֱ����п�
    Const W_ǩ�ֿ�_����_mm As Single = 30

    '������ҳ����㣨���Ŀ�ߣ�
    Dim doc As Document: Set doc = ActiveDocument
    Dim ps As PageSetup: Set ps = doc.PageSetup
    Dim ���ÿ� As Single, ���ø� As Single, �� As Single, �� As Single
    ���ÿ� = ps.PageWidth - ps.LeftMargin - ps.RightMargin
    ���ø� = ps.PageHeight - ps.TopMargin - ps.BottomMargin
    �� = ps.LeftMargin + MM(�ڱ߾�_mm)
    �� = ps.TopMargin + MM(�ڱ߾�_mm)
    ���ÿ� = ���ÿ� - 2 * MM(�ڱ߾�_mm)
    ���ø� = ���ø� - 2 * MM(�ڱ߾�_mm)

    '���ģ�����ɶ���ɾ�� DEPT��
    ɾ��������� "COVER_PROJ"
    ɾ��������� "COVER_TITLE"
    ɾ��������� "COVER_SIGNBOX"
    ɾ��������� "COVER_ORG"
    ɾ��������� "COVER_DATE"

    '���壩ȷ�����������׶�����ʽ����Ŀ������������
    Ensure_Cover_Styles

    '��������Ŀ�����ı��򣺿�=���ģ���>���У���ʽ=����-��Ŀ����
    �����ı��� tag:="COVER_PROJ", txt:=��_��Ŀ��, _
        x:=�� - 5, y:=�� + MM(Y_��Ŀ��_mm), w:=���ÿ�, h:=MM(H_��Ŀ��_mm), _
        applyStyle:="����-��Ŀ��"

    '���ߣ����������ı��򣺿�=���ģ���>���У���ʽ=����-��������
    �����ı��� tag:="COVER_TITLE", txt:=��_������, _
        x:=��, y:=�� + MM(Y_������_mm), w:=���ÿ� + MM(10), h:=MM(H_������_mm), _
        applyStyle:="����-������"

    '���ˣ�ǩ�ֱ��Է����ı�������ĵ��Ժ����ۣ�
    ����ǩ�ֱ� "COVER_SIGNBOX", y:=�� + MM(Y_ǩ�ֿ�_mm), w:=MM(70), h:=MM(40), _
        leftColWidthMM:=28, rightColWidthMM:=35, rowHeightMM:=12, _
        fontName:="����", fontPt:=14, bold:=True, _
        centerOnPage:=True


    '���ţ����ϲ����һ�Σ�
    �����ı��� tag:="COVER_ORG", txt:=��_���, _
        x:=��, y:=�� + MM(Y_���_mm), w:=���ÿ�, h:=MM(H_���_mm), _
        applyStyle:="����-���", center:=True

    '��ʮ������
    �����ı��� tag:="COVER_DATE", txt:=��_����, _
        x:=��, y:=�� + MM(Y_����_mm), w:=���ÿ�, h:=MM(H_����_mm), _
        applyStyle:="����-����", center:=True

    MsgBox "����������ɣ��°棺��ʽ����Ŀ��/������������Ѻϲ�����", vbInformation
End Sub

'===============================
'������A��ͳһ�����ı��򣨿�ѡ�ס�������ʽ����
'  - ���ṩ applyStyle������ʹ�ø���ʽ��������/�оࣩ������ fontName/pt ����
'  - ���򰴴���� fontName/fontPt/lineSpaceMultiple ����
'===============================
Private Sub �����ı���( _
    ByVal tag As String, ByVal txt As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
    Optional ByVal applyStyle As String = "", _
    Optional ByVal center As Boolean = True, _
    Optional ByVal fontName As String = "����", _
    Optional ByVal fontPt As Single = 12, _
    Optional ByVal bold As Boolean = False, _
    Optional ByVal lineSpaceMultiple As Single = 1.5)

    Dim shp As Shape
    Set shp = ActiveDocument.Shapes.AddTextBox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=x, Top:=y, width:=w, Height:=h, _
                anchor:=ActiveDocument.Range(0, 0))   ' ���Զ�λ

    With shp
        .line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .AlternativeText = tag
        .LockAnchor = True
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Left = wdShapeCenter
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Top = y

        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0:  .MarginBottom = 0
            .TextRange.text = txt

            If Len(applyStyle) > 0 Then
                On Error Resume Next
                .TextRange.Style = ActiveDocument.Styles(applyStyle)
                On Error GoTo 0
                ' ʹ����ʽʱ������֤���룻�о�����ʽ����
                .TextRange.ParagraphFormat.alignment = IIf(center, wdAlignParagraphCenter, wdAlignParagraphLeft)
            Else
                ' ֱ��������/�о�
                With .TextRange.ParagraphFormat
                    .alignment = IIf(center, wdAlignParagraphCenter, wdAlignParagraphLeft)
                    .SpaceBefore = 0: .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpaceMultiple
                    .LineSpacing = 12 * lineSpaceMultiple
                End With
                With .TextRange.Font
                    .NameFarEast = fontName
                    .NameAscii = fontName
                    .Size = fontPt
                    .bold = bold
                End With
            End If
        End With
    End With
End Sub

'===============================
'������B������ǩ�ֱ�������ԭʵ�֣�
'===============================
'===============================
' ����ǩ�ֱ��ı�����Ƕ���
'  - centerOnPage=True���ı�����ԡ���ҳ��ˮƽ+��ֱ����
'  - ���壺���塢�Ӵ֡��ĺţ�fontPt=14��
'  - �оࣺ1.5���������оࣩ
'===============================
Private Sub ����ǩ�ֱ�( _
        ByVal tag As String, _
        Optional ByVal x As Single = 0, _
        Optional ByVal y As Single = 0, _
        Optional ByVal w As Single = 220, _
        Optional ByVal h As Single = 30, _
        Optional ByVal leftColWidthMM As Single = 30, _
        Optional ByVal rightColWidthMM As Single = 0, _
        Optional ByVal rowHeightMM As Single = 0, _
        Optional ByVal fontName As String = "����", _
        Optional ByVal fontPt As Single = 14, _
        Optional ByVal bold As Boolean = True, _
        Optional ByVal centerOnPage As Boolean = True _
        )
    
    Dim doc As Document: Set doc = ActiveDocument
    Dim ps As PageSetup: Set ps = Selection.Range.Sections(1).PageSetup
    
    
    ' �����������ر����ı��򣨾��Զ�λ��ҳ�����꣩
    Dim shp As Shape
    Set shp = doc.Shapes.AddTextBox(msoTextOrientationHorizontal, _
        x, y, w, h, doc.Range(0, 0))
    With shp
        .AlternativeText = tag
        .line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .LockAnchor = True
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Top = y
    End With
    
    If centerOnPage Then
        shp.Left = wdShapeCenter        ' ����ڡ�ҳ�桱����
    Else
        shp.Left = x                    ' ��Ҫʱ�Կ��þ��� x
    End If
    
    ' �������ı����ڲ����� 3��2 ���
    shp.TextFrame.TextRange.text = ""
    shp.TextFrame.TextRange.Select
    Dim tb As Table
    Dim wPt As Single: wPt = w
    Dim leftPt As Single: leftPt = MM(leftColWidthMM)
    Dim rightPt As Single
    
    
    '���ұ߾�Ĳ���������
    If rightColWidthMM > 0 Then
        rightPt = MM(rightColWidthMM)
    Else
        rightPt = wPt - leftPt
    End If
    
    If leftPt + rightPt > wPt Then
        rightPt = wPt - leftPt
        If rightPt < 1 Then rightPt = 1
    End If
    
    
    Set tb = doc.Tables.Add(Selection.Range, 3, 2)
    
    With tb
        .AllowAutoFit = False
        .Borders.enable = False
        .rows.alignment = wdAlignRowCenter
        .rows.AllowBreakAcrossPages = False
        .TopPadding = 0: .BottomPadding = 0
        .LeftPadding = 0: .RightPadding = 0
        .Range.Cells.VerticalAlignment = wdCellAlignVerticalBottom
        If rowHeightMM > 0 Then
            .rows.HeightRule = wdRowHeightExactly
            .rows.Height = MM(rowHeightMM)
        End If
        
        ' �п����й̶�������ռ��ʣ����
        .Columns(1).width = leftPt
        .Columns(2).width = rightPt
        
        ' �����ı�
        .cell(1, 1).Range.text = "��  �ƣ�"
        .cell(2, 1).Range.text = "��  �ˣ�"
        .cell(3, 1).Range.text = "��  ����"
        
        Dim r As Long
        For r = 1 To 3
            '����-1������ǩ���ߣ����ױߣ�**�Ӵ�**
            With .cell(r, 2).Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt   ' �� �Ӵ�
                .Color = wdColorAutomatic
            End With
        
            '����-2���ֱ�������ʽ����=�Ҷ��룻��=���С��κ�0��1.5���оࣩ
            .cell(r, 1).Range.Style = ActiveDocument.Styles("����-ǩ����")
            .cell(r, 2).Range.Style = ActiveDocument.Styles("����-ǩ����")
        Next r
    End With
End Sub

' mm �� pt
Private Function MM(mmVal As Single) As Single
    MM = mmVal * 2.835
End Function


'===============================
'������C������/�������׶���������ʽ
'   - ����-��Ŀ��������/���壬Сһ(24pt)��1.5���о࣬����
'   - ����-���������������壬Ӣ��Times New Roman��С��(36pt)��1.5���о࣬����
'===============================
Private Sub Ensure_Cover_Styles()
    Call EnsureParagraphStyle( _
        styleName:="����-��Ŀ��", _
        nameCN:="����", nameEN:="����", _
        ptSize:=24, isBold:=False, lineRule:=wdLineSpaceSingle, align:=wdAlignParagraphCenter)

    Call EnsureParagraphStyle( _
        styleName:="����-������", _
        nameCN:="����", nameEN:="Times New Roman", _
        ptSize:=36, isBold:=True, lineRule:=wdLineSpaceSingle, align:=wdAlignParagraphCenter)
        
        
    EnsureParagraphStyle "����-���", "����", "Times New Roman", 14, True, wdLineSpace1pt5, wdAlignParagraphCenter
    EnsureParagraphStyle "����-����", "����", "Times New Roman", 16, True, wdLineSpace1pt5, wdAlignParagraphCenter
    EnsureParagraphStyle "����-ǩ����", "����", "Times New Roman", 14, True, SIGN_MIN_LINESPACE_PT, wdAlignParagraphRight
    EnsureParagraphStyle "����-ǩ����", "����", "Times New Roman", 15, False, SIGN_MIN_LINESPACE_PT, wdAlignParagraphCenter
    
    '�����У�ȥ���ױ߾ࣻ����ͳһ 1.5 ���о࣬��ǰ�����㣩
    With ActiveDocument.Styles("����-ǩ����").ParagraphFormat
        .SpaceBefore = 0: .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = SIGN_MIN_LINESPACE_PT
    End With
    
    With ActiveDocument.Styles("����-ǩ����").ParagraphFormat
        .SpaceBefore = 0: .SpaceAfter = 0   ' �� ����ȥ���ױ߾�
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = SIGN_MIN_LINESPACE_PT
    End With
    

End Sub


'===============================
'������E���� TAG ɾ���ɶ���
'===============================
Private Sub ɾ���������(tag As String)
    Dim s As Shape, i As Long
    On Error Resume Next
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        Set s = ActiveDocument.Shapes(i)
        If LCase$(s.AlternativeText) = LCase$(tag) Then s.Delete
    Next i
    On Error GoTo 0
End Sub


