Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Param1 is required"
  End If
  ' ... rest of the function
End Function