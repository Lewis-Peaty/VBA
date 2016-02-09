

Function GroupBy(Groupme, modu, NumDigits)

If IsNull(Groupme) Then
GroupBy = "unknown"
ElseIf Not IsNumeric(Groupme) Then
GroupBy = "unknown"
Else
GroupBy = GenerateBucketString(Groupme,modu,NumDigits)
End If

End Function


Function padzeros(WorkingNo, Padto) As String
Dim i As Integer
Dim WorkingString As String
If IsNull(WorkingNo) Then
padzeros = ""
Else
WorkingString = WorkingNo

For i = 1 To Padto - Len(WorkingString)
padzeros = 0 & padzeros
Next i

padzeros = padzeros & WorkingString
End If
End Function

Function GenerateBucketString(Groupme, modu, NumDigits) As String
    Dim Firsthalf, Secondhalf As String
    Dim Rounded As Long
    Rounded = Int(Groupme) 'Int() added by me- prevents "Groupme < Firsthalf"
    Firsthalf = padzeros((Rounded \ modu) * modu, NumDigits)
    Secondhalf = padzeros((Rounded \ modu + 1) * modu - 1, NumDigits)
    GenerateBucketString = Firsthalf & "-" & Secondhalf
End Function
