Sub ApplyTableStyle()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Style = "Medium Shading 1 - Accent 6"
        tbl.Rows.Alignment = wdAlignRowCenter
        tbl.AutoFitBehavior wdAutoFitWindow
    Next
    For Each cht In ActiveDocument.InlineShapes
        cht.Height = 300
        cht.Width = 500
    Next
End Sub

