Attribute VB_Name = "func_printShift"
Public Function printShift(shiftTime As Variant, shiftRole As Variant) As String

Dim tempStr As String
tempStr = timeString(shiftTime)

printShift = shiftRole & " " & tempStr

End Function
