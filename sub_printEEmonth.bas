Attribute VB_Name = "sub_printEEmonth"
Public Sub printEEmonth(eeNum As Integer)


Dim wb As Workbook
Set wb = ThisWorkbook

Dim ds As Worksheet
Dim es As Worksheet
Dim ms As Worksheet

Set ds = wb.Sheets("Data")
Set es = wb.Sheets("Entry")
Set ms = wb.Sheets("MonthSchedule")

Dim monthName As String
Dim monthNum As Integer
Dim year As String
Dim eeName As String
Dim firstDate As Date

monthName = ds.Range("C4").Value
monthNum = ds.Range("B4").Value
year = ds.Range("B5").Value
eeName = ds.Range(Cells(eeNum + 22, 31).Address).Value
firstDate = ds.Range("E4").Value

Dim startRow As Integer
Dim endRow As Integer
startRow = ds.Range("F4").Value
endRow = ds.Range("F5").Value

Dim eeData As Variant
Dim monthArray(0 To 4, 0 To 5) As Variant

ms.Range("A1") = eeName
ms.Range("D1") = monthName
ms.Range("E1") = year
ms.Range("A3") = firstDate

Dim weekIndex As Integer
Dim dayIndex As Integer
Dim rowIndex As Integer


For weekIndex = 0 To 4
    For dayIndex = 0 To 5
        eeData = es.Range(Cells(startRow + weekIndex * 49 + dayIndex * 7, 3 + eeNum * 4).Address & ":" & _
                            Cells(startRow + weekIndex * 49 + dayIndex * 7 + 6, 4 + eeNum * 4).Address)
        monthArray(weekIndex, dayIndex) = printDay(eeData)
    Next dayIndex
Next weekIndex

For weekIndex = 0 To 4
    ms.Range(Cells(weekIndex * 2 + 4, 1).Address) = monthArray(weekIndex, 0)
    ms.Range(Cells(weekIndex * 2 + 4, 2).Address) = monthArray(weekIndex, 1)
    ms.Range(Cells(weekIndex * 2 + 4, 3).Address) = monthArray(weekIndex, 2)
    ms.Range(Cells(weekIndex * 2 + 4, 4).Address) = monthArray(weekIndex, 3)
    ms.Range(Cells(weekIndex * 2 + 4, 5).Address) = monthArray(weekIndex, 4)
    ms.Range(Cells(weekIndex * 2 + 4, 6).Address) = monthArray(weekIndex, 5)
Next weekIndex



End Sub
