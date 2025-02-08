Function MyFunction(param1)
  On Error GoTo ErrHandler
  If IsEmpty(param1) Then
    Err.Raise 13, , "Param1 is required"
  End If
  ' ... rest of the function
  Exit Function
ErrHandler:
  MsgBox "An error occurred: " & Err.Description, vbCritical
  'Further error handling or logging
End Function