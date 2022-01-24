Attribute VB_Name = "func_timeString"
Public Function timeString(timeSTR As Variant) As String

Dim inTime As Double
Dim outTime As Double

inTime = CDbl(Left(timeSTR, InStr(timeSTR, "-") - 1))
outTime = CDbl(Right(timeSTR, Len(timeSTR) - InStr(timeSTR, "-")))

If inTime > 12 Then
    inTime = inTime - 12
End If

If outTime > 12 Then
    outTime = outTime - 12
End If

Dim inString As String

Dim outString As String

If inTime <> Int(inTime) Then
    inString = Int(inTime) & ":" & ((inTime - Int(inTime)) * 60)
Else
    inString = inTime & ":00"
End If

If outTime <> Int(outTime) Then
    outString = Int(outTime) & ":" & ((outTime - Int(outTime)) * 60)
Else
    outString = outTime & ":00"
End If

timeString = inString & "-" & outString

End Function
