Public Sub emptyClipboard()
    Application.CutCopyMode = False
End Sub

Public Sub Copy_Visible()
    Dim myRange As Range
    Set myRange = Selection.SpecialCells(xlCellTypeVisible)
    myRange.Copy
End Sub

Public Sub ClipCommaArray()
    Dim myRange As Range
    Dim myArray()
    Dim tempstring As String, arraystring As String
    
    Set myRange = Selection.SpecialCells(xlCellTypeVisible)
    myArray = Tools.rangeToArray(myRange)
    arraystring = Join(myArray, ",")
    Call Tools.ClipboardThis(arraystring)
    
    MsgBox ("Array is now copied to clipboard.")

End Sub
