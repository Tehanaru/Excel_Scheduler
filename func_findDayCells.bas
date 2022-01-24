Attribute VB_Name = "func_findDayCells"
Public Function findDayCells(dateRow As Integer, eeNum As Integer) As Variant

Dim wb As Workbook
Set wb = ThisWorkbook

Dim es As Worksheet
Set es = wb.Sheets("Entry")

Dim eeCol As Integer
eeCol = eeNum * 4 + 3

Set findDayCells = es.Range(Cells(dateRow, eeCol).Address & ":" & _
                            Cells(dateRow + 4, eeCol + 1).Address)


End Function


