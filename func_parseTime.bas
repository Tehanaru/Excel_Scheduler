Attribute VB_Name = "func_parseTime"
Public Function parseTime(timeString As String) As Double

If Len(timeString) < 2 Then
    parseTime = 0
    Exit Function
End If

Dim tIn As Double
Dim tOut As Double

tIn = CDbl(Left(timeString, InStr(1, timeString, "-") - 1))
tOut = CDbl(Right(timeString, Len(timeString) - InStr(1, timeString, "-")))

parseTime = tOut - tIn

End Function
