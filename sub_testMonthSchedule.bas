Attribute VB_Name = "sub_testMonthSchedule"
Public Sub testMonthSchedule()

Dim wb As Workbook
Set wb = ThisWorkbook

Dim ms As Worksheet
Set ms = wb.Sheets("MonthSchedule")

Call printEEmonth(ms.Range("A13").Value)


End Sub
