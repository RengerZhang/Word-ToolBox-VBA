Attribute VB_Name = "�ĵ��ָ�"
Sub SaveSelected()
'UpdatebyExtendoffice20181115
    Selection.Copy
    Documents.Add , , wdNewBlankDocument
    Selection.Paste
    ActiveDocument.Save
    'ActiveDocument.Close
End Sub
