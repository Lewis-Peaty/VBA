Sub update_ranges()
    
    Dim str As String
    str = "C9:C" & Worksheets("Breakdown List").Range("C9").End(xlDown).Row
    Worksheets("Breakdown List").Range(str).Name = "Breakdown_Defect"
    
    str = "E9:E" & Worksheets("Breakdown List").Range("E9").End(xlDown).Row
    Worksheets("Breakdown List").Range(str).Name = "Breakdown_Performance"
    
    str = "A9:A" & Worksheets("Breakdown List").Range("A9").End(xlDown).Row
    Worksheets("Breakdown List").Range(str).Name = "Breakdown_Profile"

End Sub
