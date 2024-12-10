Function f(a, b)
  If IsEmpty(a) Or IsNumeric(a) = False Then
    MsgBox "Error: Invalid argument 'a'", vbCritical
    Exit Function
  End If
  If IsEmpty(b) Or IsNumeric(b) = False Then
    MsgBox "Error: Invalid argument 'b'", vbCritical
    Exit Function
  End If
  f = CDbl(a) + CDbl(b) 
End Function

MsgBox f(1, "")
MsgBox f(1, 2)
MsgBox f("", 2)