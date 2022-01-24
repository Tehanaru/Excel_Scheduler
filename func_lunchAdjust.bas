Attribute VB_Name = "func_lunchAdjust"
Public Function lunchAdjust(shiftArray As Variant) As Integer

lunchAdjust = 0

If shiftArray(1, 1) = 0 Then
    lunchAdjust = 0
    Exit Function
End If

Dim timeArray(1 To 5, 1 To 2) As Double
Dim i As Integer

For i = 1 To 5
    timeArray(i, 1) = parseTime(CStr(shiftArray(i, 1)))
    timeArray(i, 2) = isClocked(CStr(shiftArray(i, 2)))
Next i


For i = 1 To 5
    If timeArray(i, 1) > 5.5 Then
        lunchAdjust = i
        Exit Function
    End If
Next i

Dim runTotal As Double
runTotal = 0

For i = 1 To 5
    If timeArray(i, 2) = 1 Then
        runTotal = runTotal + timeArray(i, 1)
    Else
        runTotal = 0
    End If

    If runTotal > 5.5 Then
        lunchAdjust = i
        Exit Function
    End If
Next i

End Function
