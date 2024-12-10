Function f(a, b)
  If IsEmpty(a) Then Exit Function
  If IsEmpty(b) Then Exit Function
  f = a + b
End Function

MsgBox f(1, "")