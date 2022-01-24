Attribute VB_Name = "sub_pasteSchedule"
Public Sub pasteSchedule()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim wb As Workbook
Set wb = ThisWorkbook

Dim sf As Worksheet
Dim es As Worksheet

Set sf = wb.Sheets("Schedule Filler")
Set es = wb.Sheets("Entry")

Dim startRow As Integer
Dim endRow As Integer
Dim eeCol As Integer

Dim overWrite As Integer


overWrite = 0

If sf.Range("J8").Value = "Yes" Then
    overWrite = overWrite + 1
End If

If sf.Range("J10").Value = "Yes" And overWrite = 1 Then
    overWrite = overWrite + 1
End If


Dim copiedRange As Variant
Set copiedRange = sf.Range("C2:F50")

Dim weekDex As Integer
Dim dayDex As Integer
Dim numWeeks As Integer
Dim rowDex As Integer

startRow = sf.Range("L14").Value
endRow = sf.Range("L18").Value
eeCol = (sf.Range("H2").Value - 1) * 4 + 7

numWeeks = (endRow - startRow + 7) / 49

Dim curRow As Integer

For weekDex = 0 To numWeeks - 1
    For dayDex = 0 To 5
        For rowDex = 0 To 6
        If rowDex < 5 Then
            curRow = startRow + rowDex + dayDex * 7 + weekDex * 49
            If es.Cells(curRow, eeCol + 1).Value = "HOL" Then
                Exit For
            End If
            Select Case overWrite
                Case 0
                    If es.Cells(curRow, eeCol).Value = 0 Then
                        es.Cells(curRow, eeCol).Value = copiedRange(dayDex * 7 + rowDex + 1, 1)
                        es.Cells(curRow, eeCol + 1).Value = copiedRange(dayDex * 7 + rowDex + 1, 2)
                    End If
                Case 1
                    If es.Cells(curRow, eeCol + 1).Value <> "PTO" Then
                        es.Cells(curRow, eeCol).Value = copiedRange(dayDex * 7 + rowDex + 1, 1)
                        es.Cells(curRow, eeCol + 1).Value = copiedRange(dayDex * 7 + rowDex + 1, 2)
                    End If
                Case 2
                    es.Cells(curRow, eeCol).Value = copiedRange(dayDex * 7 + rowDex + 1, 1)
                    es.Cells(curRow, eeCol + 1).Value = copiedRange(dayDex * 7 + rowDex + 1, 2)
            End Select
        End If
        Next rowDex
    Next dayDex
Next weekDex


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub
