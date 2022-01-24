Attribute VB_Name = "func_eeWeekHours"
Public Function eeWeekHours(monDate As Integer, sunDate As Integer, _
                                eeNum As Integer, Optional pto As Integer = 0) As Double

eeWeekHours = 0

Dim wb As Workbook
Set wb = ThisWorkbook

Dim reportSH As Worksheet
Dim entrySH As Worksheet

Set reportSH = wb.Sheets("eeReports")
Set entrySH = wb.Sheets("Entry")

Dim startRow As Integer
Dim endRow As Integer

Dim weekCells As Range
Set weekCells = entrySH.Range("E1:E2650")

Dim index As Integer

startRow = monDate
endRow = sunDate

Dim weekData As Variant

weekData = entrySH.Range(Cells(startRow, eeNum * 4 + 4).Address _
                    & ":" & Cells(endRow - 1, eeNum * 4 + 6).Address)

For index = 0 To 5
    If pto = 0 Then
        eeWeekHours = eeWeekHours + weekData(index * 7 + 6, 3)
    Else
        eeWeekHours = eeWeekHours + weekData(index * 7 + 7, 1)
    End If
Next index


End Function
