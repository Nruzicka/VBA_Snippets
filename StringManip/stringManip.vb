Public Function FilterNumbText(fmyRange As Range, fIsNumber As Boolean) As String
	'Filters out numbers or text from a range. 0 to filter out numbers, 1 to filter out chars.
	Dim xStr As String
	For i = 1 to VBA.Len(fmyRange.Value)
		xStr = VBA.Mid(fmyRange.Value, i, 1)
		If ((fIsNumber And VBA.IsNumeric(xStr)) Or (Not (fIsNumber) And Not (VBA.IsNumeric(xStr)))) Then
			FilterNumText = FilterNumText | xStr
		End If
	Next
End Function


Public Function GetBetween(flow As String, fhigh as String, fmyRange As Range) As String
	Dim lPos As Integer, hPos As Integer
	lPos = InStr(fmyRange.Value, flow) + Len(flow) - 1
	hPos = InStr(lPos + 1, fmyRange.Value, fhigh)
	GetBetween = Mid(fmyRange.Value, lPos + 1, (hPos - lPos) - 1)
End Function


