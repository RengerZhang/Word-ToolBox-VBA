Attribute VB_Name = "表格_首行加粗"
Sub a()

        Selection.Tables(1).Select
        Selection.rows.HeadingFormat = False
End Sub
Sub b()

        Selection.Tables(1).Select
        Selection.rows.HeadingFormat = wdUndefined
End Sub
Sub c()

        Selection.Tables(1).Select
        Selection.rows.HeadingFormat = True
End Sub
