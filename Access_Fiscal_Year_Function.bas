Option Compare Database
Const FMonthStart = 7   ' Numeric value representing the first month
                        ' of the fiscal year.
Const FDayStart = 1     ' Numeric value representing the first day of
                        ' the fiscal year.
Const FYearOffset = 0   ' 0 means the fiscal year starts in the
                        ' current calendar year.
                        ' -1 means the fiscal year starts in the
                        ' previous calendar year.
                        
Function GetFiscalYear(ByVal x As Variant)
   If x < DateSerial(Year(x), FMonthStart, FDayStart) Then
      GetFiscalYear = (Year(x) - 1) & "/" & Year(x) '- FYearOffset - 1
   Else
      GetFiscalYear = Year(x) & "/" & Year(x) + 1 '- FYearOffset
   End If
End Function



-- VARIANT WITH ONLY 2 DIGIT FY ---

Const FMonthStart = 7   ' Numeric value representing the first month
                        ' of the fiscal year.
Const FDayStart = 1     ' Numeric value representing the first day of
                        ' the fiscal year.
Const FYearOffset = 0   ' 0 means the fiscal year starts in the
                        ' current calendar year.
                        ' -1 means the fiscal year starts in the
                        ' previous calendar year.
                        
Function GetFiscalYear(ByVal x As Variant)
   If x < DateSerial(Year(x), FMonthStart, FDayStart) Then
      GetFiscalYear = ((Year(x) - 2000) - 1) & "/" & (Year(x) - 2000) '- FYearOffset - 1
   Else
      GetFiscalYear = (Year(x) - 2000) & "/" & (Year(x) - 2000) + 1 '- FYearOffset
   End If
End Function
