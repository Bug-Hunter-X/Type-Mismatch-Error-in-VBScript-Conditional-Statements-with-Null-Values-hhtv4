Function f(a,b)
  If a>b then
    MsgBox "a is greater than b"
  ElseIf a<b then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End if
end function

'The error occurs when the comparison involves null values.
'If 'a' or 'b' is null then the comparison may not work as expected
'Resulting in a type mismatch error.