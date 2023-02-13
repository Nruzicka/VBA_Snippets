Private Function toAmerDate(target, delimiter)
    'converts dd/mm/yy format to mm/dd/yy
    Dim splitarray() As String
    Dim month As String, day As String, year As String
    
    splitarray = Split(target, delimiter)
    day = splitarray(0)
    month = splitarray(1)
    year = splitarray(2)
    
    toAmerDate = month & "/" & day & "/" & year
End Function

Private Function dateDotRemove(target)
    dateDotRemove = Replace(target.Text, ".", "/")
End Function

Public Function plxsWorkday(DateRange As Range, Optional style As Single) As Variant
    'This function offsets any weekends and holidays. Set style to 1 if you are using dd/mm/yy format.
    Dim moddate As Date
    Dim amerdate As String, tempstring As String
    tempstring = dateDotRemove(DateRange)
    
    'An array of US holidays to look up.
    holidays = Array("12/24/2019", "12/25/2019", "1/1/2020", "5/25/2020", "7/3/2020", "7/4/2020", "9/7/2020", "11/26/2020", "11/27/2020", "12/24/2020", "12/25/2020", "12/26/2020", _
    "1/1/2021", "1/2/2021", "4/5/2021", "5/31/2021", "7/4/2021", "9/6/2021", "11/25/2021", "11/26/2021", "12/24/2021", "12/25/2021", "12/26/2021", _
    "12/31/2021", "1/1/2022", "1/3/2022", "5/30/2022", "7/4/2022", "7/5/2022", "9/5/2022", "11/24/2022", "11/25/2022", "12/23/2022", "12/25/2022", _
    "12/26/2022", "1/1/2023", "1/2/2023", "5/29/2023", "7/4/2023", "9/4/2023", "11/23/2023", "11/24/2023", "12/24/2023", "12/25/2023", "12/26/2023", _
    "1/1/2024", "5/27/2024", "7/4/2024", "9/2/2024", "11/28/2024", "11/29/2024", "12/23/2024", "12/24/2024", "12/25/2024", _
    "1/1/2025", "5/26/2025", "7/4/2025", "9/1/2025", "11/27/2025", "11/28/2025", "12/24/2025", "12/25/2025", "12/26/2025", _
    "1/1/2026", "1/2/2026")
    
    'checking if dates are in dd/mm/yy format
    If style = 1 Then
        amerdate = toAmerDate(tempstring, "/")
        'convert string to date and move off holidays.
        moddate = Application.WorksheetFunction.WorkDay(CDate(amerdate) - 1, 1, holidays)
     Else
        moddate = Application.WorksheetFunction.WorkDay(CDate(tempstring) - 1, 1, holidays)
    End If
    
    plxsWorkday = Format(moddate, "MM/DD/YYYY")
End Function

Public Function PlxsQtr(DateRange As Range) As Variant 'Returns the fiscal quarter that the date falls in. CAUTION 
    Dim tempstring As String, moddate As Date, qtrdate As Date
    tempstring = dateDotRemove(DateRange)
    
    'converting back to date
    moddate = CDate(tempstring)
    
    ' bucketing dates into quarters
    If Format(moddate, "m") = 1 Or Format(moddate, "m") = 2 Or Format(moddate, "m") = 3 Then
        qtrdate = "1 / 15 / " & year(moddate)
        PlxsQtr = Format(qtrdate, "mmm yy")
    ElseIf Format(moddate, "m") = 4 Or Format(moddate, "m") = 5 Or Format(moddate, "m") = 6 Then
        qtrdate = "4/15/" & year(moddate)
        PlxsQtr = Format(qtrdate, "mmm yy")
    ElseIf Format(moddate, "m") = 7 Or Format(moddate, "m") = 8 Or Format(moddate, "m") = 9 Then
        qtrdate = "7/15/" & year(moddate)
        PlxsQtr = Format(qtrdate, "mmm yy")
    ElseIf Format(moddate, "m") = 10 Or Format(moddate, "m") = 11 Or Format(moddate, "m") = 12 Then
        qtrdate = "10/15/" & year(moddate)
        PlxsQtr = Format(qtrdate, "mmm yy")
   End If
End Function

Function weekDate(dateRange As Date, Optional dayNum As Single = 1) As Date
    Dim newYearDay As Date, adjDate As Date, outDate As Date
    Dim NYWeekDay As Integer, weekNum As Integer, adjWeekNum As Integer
    
    newYearDay = CDate("1/1/" & year(dateRange))
    NYWeekDay = Application.WorksheetFunction.Weekday(newYearDay, dayNum)  'TODO: test modifying the default value of daynum to toggle monday, friday, etc
    weekNum = Application.WorksheetFunction.weekNum(dateRange) - 1
    adjDate = newYearDay - NYWeekDay
    adjWeekNum = weekNum * 7
    outDate = adjDate + 1 + adjWeekNum
    

    weekDate = Application.WorksheetFunction.Max(newYearDay, outDate) 'TODO use the format function to make it mm/dd/yyyy
End Function
