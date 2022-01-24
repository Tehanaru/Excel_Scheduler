Attribute VB_Name = "func_printDay"
Public Function printDay(dayCells As Variant) As String

Dim numShifts As Integer

Dim i As Integer

numShifts = 0

For i = 1 To 5
    If Len(dayCells(i, 1)) > 1 Then
        numShifts = numShifts + 1
    End If
Next i

If numShifts = 0 Then
    printDay = ""
    Exit Function
End If


Select Case numShifts
    Case 1
        printDay = printShift(dayCells(1, 1), dayCells(1, 2))
    Case 2
        printDay = printShift(dayCells(1, 1), dayCells(1, 2)) _
        & vbNewLine & printShift(dayCells(2, 1), dayCells(2, 2))
    Case 3
        printDay = printShift(dayCells(1, 1), dayCells(1, 2)) _
        & vbNewLine & printShift(dayCells(2, 1), dayCells(2, 2)) _
        & vbNewLine & printShift(dayCells(3, 1), dayCells(3, 2))
    Case 4
        printDay = printShift(dayCells(1, 1), dayCells(1, 2)) _
        & vbNewLine & printShift(dayCells(2, 1), dayCells(2, 2)) _
        & vbNewLine & printShift(dayCells(3, 1), dayCells(3, 2)) _
        & vbNewLine & printShift(dayCells(4, 1), dayCells(4, 2))
    Case 5
        printDay = printShift(dayCells(1, 1), dayCells(1, 1)) _
        & vbNewLine & printShift(dayCells(2, 1), dayCells(2, 2)) _
        & vbNewLine & printShift(dayCells(3, 1), dayCells(3, 2)) _
        & vbNewLine & printShift(dayCells(4, 1), dayCells(4, 2)) _
        & vbNewLine & printShift(dayCells(5, 1), dayCells(5, 2))
End Select

        



End Function
