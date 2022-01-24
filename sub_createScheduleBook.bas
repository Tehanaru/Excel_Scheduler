Attribute VB_Name = "sub_createScheduleBook"
Public Sub createScheduleBook()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
On Error Resume Next

Dim schedB As Workbook
Dim printB As Workbook

Set schedB = ThisWorkbook
Set printB = Workbooks.Add

Dim monthS As Worksheet
Dim dataS As Worksheet
Dim entryS As Worksheet

Set monthS = schedB.Sheets("MonthSchedule")
Set dataS = schedB.Sheets("Data")

Dim monthName As String
Dim year As String

monthName = dataS.Range("C4").Value
year = dataS.Range("B5").Value

Dim numEEs As Integer
Dim inclSep As Integer
inclSep = 0
numEEs = dataS.Range("B3")
monthNum = dataS.Range("B4")

If dataS.Range("D7").Value = "Yes" Then
    inclSep = 1
End If

Dim eeStatus As Integer

Dim index As Integer
For index = numEEs To 1 Step -1
    eeStatus = dataS.Range(Cells(22 + index, 27).Address).Value
    If inclSep = 1 Or eeStatus = -1 Or eeStatus >= monthNum Then
        Call printEEmonth(index)
        monthS.Copy before:=printB.Sheets(1)
        printB.Sheets(1).Name = dataS.Range(Cells(22 + index, 31).Address).Value
        With printB.Sheets(1).UsedRange
            .Value = .Value
        End With
    Else
        'placeholder
    End If
Next index

Application.DisplayAlerts = False

printB.Sheets("Sheet1").Delete
printB.SaveAs "H:\Schedules\Front Office Schedule " & monthName & " " & year & ".xlsx"

Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
