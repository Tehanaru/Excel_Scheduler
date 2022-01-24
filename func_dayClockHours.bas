Attribute VB_Name = "func_dayClockHours"
Public Function dayClockHours(dateRow As Integer, _
                                Optional allHours As Integer = 0) As Double

dayClockHours = 0

Dim wb As Workbook
Set wb = ThisWorkbook

Dim es As Worksheet
Set es = wb.Sheets("Entry")

Dim dayData As Variant

dayData = es.Range(Cells(dateRow, 7).Address & ":" & Cells(dateRow + 6, 86).Address)

Dim index As Integer

If allHours <> 0 Then
    For index = 1 To 30
        dayClockHours = dayClockHours + dayData(6, index * 4)
    Next index
    Exit Function
Else
    dayClockHours = dayClockHours + roleHours(dayData, "MFD")
    dayClockHours = dayClockHours + roleHours(dayData, "MCC")
    dayClockHours = dayClockHours + roleHours(dayData, "DFD")
    dayClockHours = dayClockHours + roleHours(dayData, "DCC")
    dayClockHours = dayClockHours + roleHours(dayData, "CLS")
    dayClockHours = dayClockHours + roleHours(dayData, "ADM")
End If

End Function
