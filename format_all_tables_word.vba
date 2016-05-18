Sub ApplyTableStyle()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Style = "Medium Shading 1 - Accent 6"
        tbl.Rows.Alignment = wdAlignRowCenter
        tbl.AutoFitBehavior wdAutoFitWindow
        'For Each c In tbl.Columns
        '    c.ParagraphFormat.Alignment = wdAlignParagraphCenter
        'Next c
    Next
End Sub
