Attribute VB_Name = "func_spliceArray5x2frm5xN"
Public Function spliceA5x2frm5xN(ByRef outArray As Variant, inArray As Variant, _
    rLow As Integer, rHigh As Integer) As Integer

spliceA5x2frm5xN = 0

Dim i As Integer

For i = 1 To 5
    outArray(i, 1) = inArray(rLow + i - 1, 1)
    outArray(i, 2) = inArray(rLow + i - 1, 2)
Next i

spliceA5x2frm5xN = 1

End Function
