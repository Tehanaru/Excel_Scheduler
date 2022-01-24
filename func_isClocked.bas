Attribute VB_Name = "func_isClocked"
Public Function isClocked(role As String) As Integer

Select Case role
    Case "MFD", "DFD", "MCC", "DCC", "EVR", "ADM", "CLS", "REM", "MMC", "SUP"
        isClocked = 1
    Case "PTO", "OUT", "HOL", "FML", "UPT"
        isClocked = 0
    Case Else
        isClocked = 0
End Select

End Function
