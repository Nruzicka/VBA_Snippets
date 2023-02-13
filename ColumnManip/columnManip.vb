Public Function COLUMNLETTER(choice As Range) As String
    COLUMNLETTER = Split(choice.Address, "$")(1)
End Function

Public Function FULLCOLUMN(choice As Range) As String
    FULLCOLUMN = Split(choice.Address, "$")(1) & ":" & Split(choice.Address, "$")(1)
End Function
