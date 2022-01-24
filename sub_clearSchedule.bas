Attribute VB_Name = "sub_clearSchedule"
Public Sub clearSchedule()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim wb As Workbook
Set wb = ThisWorkbook

Dim ws As Worksheet
Set ws = wb.ActiveSheet

Dim mon As Range
Dim tue As Range
Dim wed As Range
Dim thu As Range
Dim fri As Range
Dim sat As Range
Dim sun As Range

Set mon = ws.Range("C2:D6")
Set tue = ws.Range("C9:D13")
Set wed = ws.Range("C16:D20")
Set thu = ws.Range("C23:D27")
Set fri = ws.Range("C30:D34")
Set sat = ws.Range("C37:D41")
Set sun = ws.Range("C44:D48")

mon.ClearContents
tue.ClearContents
wed.ClearContents
thu.ClearContents
fri.ClearContents
sat.ClearContents
sun.ClearContents

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
