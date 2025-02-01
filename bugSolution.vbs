Function f(x)
  Dim result
  result = 1
  If x > 1 Then
    For i = 2 To x
      result = result * i
    Next
  End If
  f = result
End Function
MsgBox f(5) 
'Alternative iterative factorial function