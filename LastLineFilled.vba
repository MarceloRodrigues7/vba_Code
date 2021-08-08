Sub LastLineFilled()
Dim line As Integer
Dim cell As String
Dim position As String

line = Range("A1048576").End(xlUp).Row
cell = "A" + CStr(line)
position = "A2:" + cell
Range(position).Select

''Paste values -> Selection.FillDown
End Sub
