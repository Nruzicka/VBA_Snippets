Public Function clipBoardThis(Optional s As String) As String
	Dim v: v = s 'Cast to variant for 64-bit VBA support
	Dim htmlObj As Object
	Set htmlObj = CreateObject("htmlfile")
	With htmlObj.parentWindow.clipboardData
		Select Case True
			Case Len(s):	.setData "text", v
			Case Else:	Clipboard = .GetData("text")
		End Select
	End With
	Set htmlObj = Nothing
End Function

Public Sub depClipboardThis(Text As String)
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Public Function rangeToArray(target As Range) As Variant
    'converts 2 dim range to 1 dim array
    'exploits the for-each loop to capture non-continuous ranges
    Dim myArray()
    ReDim myArray(target.Count - 1)
    Dim dye As Variant
    
    Dim i As Integer
    i = 0
    For Each rngcell In target
        dye = rngcell.Value
        myArray(i) = dye
        i = i + 1
    Next rngcell
        
    rangeToArray = myArray
End Function


Sub unpivot2()
    'setting up input and output sheets
    Dim inputSheet As Worksheet, outputSheet As Worksheet
    Dim rowRng As Range, fields As Range, item As Range
    
    Set inputSheet = ActiveSheet
    ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = "Unpivot Output" & Sheets.Count
    Set outputSheet = ActiveWorkbook.Sheets(Sheets.Count)
    outputSheet.Range("A1:C1") = Array("Item", "Date", "Qty")
    inputSheet.Activate
    
    'loop unpivot to output worksheet
    Set rowRng = inputSheet.Range(Range("A2"), Range("A" & Rows.Count).End(xlUp))
    
    Dim c As Integer
    c = 2
    For Each item In rowRng
        Set fields = inputSheet.Range(Range("B" & item.Row), Cells(item.Row, Columns.Count).End(xlToLeft))
        With outputSheet
            .Cells(c, 1).Resize(fields.Count) = item
            .Cells(c, 2).Resize(fields.Count) = Application.Transpose(rowRng(1).Offset(-1, 1).Resize(, fields.Count).Value)
            .Cells(c, 3).Resize(fields.Count) = Application.Transpose(fields.Value)
        End With
        c = c + fields.Count
    Next item

End Sub


Sub resetDataType()
'Uses the text-to-columns trick to reset the data type
    Dim myRange As Range
    Dim address As String, destination As String
    Set myRange = Selection
    address = strQuote & Replace(myRange.address, "$", "") & strQuote
    
    myRange.TextToColumns Range(address), xlDelimited, xlTextQualifierDoubleQuote, False, True, False, False, False, False, , Array(1, 1), , , True
End Sub
