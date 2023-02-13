Public Function COLUMNLETTER(choice As Range) As String
    COLUMNLETTER = Split(choice.Address, "$")(1)
End Function

Public Function FULLCOLUMN(choice As Range) As String
    FULLCOLUMN = Split(choice.Address, "$")(1) & ":" & Split(choice.Address, "$")(1)
End Function

public Function colToRange(mySheet As Worksheet, _
	startCol As String, lastCol As String, _
	startIndex As String, Optional refCol As String = vbNullString) As Range

	Dim myRange As Range
	Dim colHeader As String

	If (refCol = vbNullString) Then
		refCol = startCol
	End If

	colHeader = startCol & startIndex & ":" & lastCol
	Set myRange = mySheet.range(colHeader & range(refCol & Rows.Count).End(xlUp).Row)

	Set colToRange = myRange
End Function
