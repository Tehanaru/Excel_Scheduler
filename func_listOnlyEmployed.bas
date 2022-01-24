Attribute VB_Name = "func_listOnlyEmployed"
Public Function listOnlyEmployed(monthNumber As Integer, eeNum As Integer, nameList As Range, sepList As Range) As String

Dim finalNameList(1 To 10) As String

'Dim nameList As Variant
'Dim sepList As Variant

'nameList = nameListData
'sepList = sepListData

Dim i As Integer
Dim n As Integer
n = 1
For i = 1 To 30
    If sepList(i) = -1 Or sepList(i) >= monthNumber Then
        finalNameList(n) = nameList(i)
        n = n + 1
    End If
Next i

If eeNum < n Then
    listOnlyEmployed = finalNameList(eeNum)
Else
    listOnlyEmployed = ""
End If


End Function
