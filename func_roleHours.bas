Attribute VB_Name = "func_roleHours"
Public Function roleHours(dayArray As Variant, role As String, Optional spanish As Integer = 0) As Double

roleHours = 0

Dim employee As Integer
Dim shift As Integer

Dim tempHrs As Double
Dim spaAdj As Integer

Dim wb As Workbook
Set wb = ThisWorkbook
Dim es As Worksheet
Dim ds As Worksheet

Set ds = wb.Sheets("Data")

Dim eeTotal As Integer
eeTotal = ds.Range("B3")

For employee = 0 To eeTotal
    tempHrs = 0
    For shift = 1 To 5
        If dayArray(1, 1 + employee * 4) = 0 Then
            Exit For
        End If
        
        If dayArray(shift, 2 + employee * 4) = role Then
            tempHrs = tempHrs + dayArray(shift, 3 + employee * 4)
        End If
    Next shift
    
    If spanish <> 0 Then
        If dayArray(7, 4 + employee * 4) <> 0 Then
            spaAdj = 1
        Else
            spaAdj = 0
        End If
    Else
        spaAdj = 1
    End If
    
    roleHours = roleHours + tempHrs * spaAdj
        
Next employee

End Function
