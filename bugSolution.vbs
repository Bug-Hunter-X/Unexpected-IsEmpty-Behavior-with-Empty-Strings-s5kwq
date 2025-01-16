Function MyFunction(param1)
  If Len(Trim(param1)) = 0 Then
    ' Handle empty or whitespace-only parameter
    Exit Function
  End If

  ' ...rest of the function...
End Function