# Opportunities First Academic Database
```
**********************************************************************
DYNAMIC CALENDAR
**********************************************************************
Private Sub Form_Activate()
CalendarDateNumbers
End Sub

Private Sub GEDCalendarMonth_Change()
CalendarDateNumbers
End Sub

Private Sub GEDCalendarYear_Change()
CalendarDateNumbers
End Sub

Private Sub NewTestApptBtn_Click()
DoCmd.OpenForm ("GEDTestScheduleEntryF")
End Sub

Private Sub NextMonth_Click()
Dim month As Integer
month = Me.GEDCalendarMonth
If month = 12 Then
    Me.GEDCalendarMonth = 1
    Me.GEDCalendarYear = Me.GEDCalendarYear + 1
Else
    Me.GEDCalendarMonth = Me.GEDCalendarMonth + 1
End If
CalendarDateNumbers
End Sub

Private Sub PreviousMonth_Click()
Dim month As Integer
month = Me.GEDCalendarMonth
If month = 1 Then
    Me.GEDCalendarMonth = 12
    Me.GEDCalendarYear = Me.GEDCalendarYear - 1
Else
    Me.GEDCalendarMonth = Me.GEDCalendarMonth - 1
End If
CalendarDateNumbers
End Sub

Public Sub CalendarDateNumbers()
' Given the current month and year selected by the user through the
' combo boxes, calculate the serial date and then the day of the week
' of the first day in that month

Dim selectedMonth As Integer
Dim selectedYear As Integer
Dim serialDate As Long

selectedMonth = Me.GEDCalendarMonth
selectedYear = Me.GEDCalendarYear
serialDate = DateSerial(selectedYear, selectedMonth, 1)
dayOfWeekFirst = Weekday(serialDate)

' Add numbers to calendar textboxes
Dim calBoxes As Variant
calBoxes = Array(Me.TextA1, Me.TextA2, Me.TextA3, Me.TextA4, Me.TextA5, Me.TextA6, Me.TextA7, _
    Me.TextB1, Me.TextB2, Me.TextB3, Me.TextB4, Me.TextB5, Me.TextB6, Me.TextB7, _
    Me.TextC1, Me.TextC2, Me.TextC3, Me.TextC4, Me.TextC5, Me.TextC6, Me.TextC7, _
    Me.TextD1, Me.TextD2, Me.TextD3, Me.TextD4, Me.TextD5, Me.TextD6, Me.TextD7, _
    Me.TextE1, Me.TextE2, Me.TextE3, Me.TextE4, Me.TextE5, Me.TextE6, Me.TextE7, _
    Me.TextF1, Me.TextF2, Me.TextF3, Me.TextF4, Me.TextF5, Me.TextF6, Me.TextF7)
    
For c = 0 To 41
    testDayBool(c) = False
Next c

Dim firstDay As String
Dim firstDayIdentified As Boolean
Dim i As Integer
firstDay = "TextA" & dayOfWeekFirst
firstDayIdenfied = False
i = 1
numDays = DaysInMonth(selectedMonth, selectedYear)
Dim dummy As Integer
Dim calBoxNum As Integer
calBoxNum = 0

For Each calBox In calBoxes
    ' Label each box with correct day of month
    If firstDayIdentified = True And i < numDays Then
        i = i + 1
        calBox.Caption = i
        testDatesInMonth(calBoxNum) = DateSerial(selectedYear, selectedMonth, i)
    Else
        calBox.Caption = ""
    End If
    If calBox.name = firstDay Then
        calBox.Caption = i
        firstDayIdentified = True
        testDatesInMonth(calBoxNum) = DateSerial(selectedYear, selectedMonth, i)
    End If
    
    ' Shade old dates and outline current date in green
    If calBox.Caption = "" Then
        calBox.Visible = False
        calBox.BorderColor = RGB(192, 192, 192)
        calBox.BorderWidth = 0
    ElseIf IsOldDate(selectedMonth, i, selectedYear) Then
        calBox.Visible = True
        calBox.BackColor = RGB(230, 230, 230)
        calBox.BorderColor = RGB(192, 192, 192)
        calBox.BorderWidth = 0
        dummy = GEDTestDateQ(selectedMonth, i, selectedYear)
    Else
        dummy = GEDTestDateQ(selectedMonth, i, selectedYear)
        calBox.Visible = True
        calBox.BackColor = RGB(255, 255, 255)
        If IsToday(selectedMonth, i, selectedYear) Then
            calBox.BorderColor = RGB(0, 204, 0)
            calBox.BorderWidth = 3
        Else
            calBox.BorderColor = RGB(192, 192, 192)
            calBox.BorderWidth = 0
        End If
    End If
    
    ' If at least one test is scheduled for that day, write on
    ' caption and link to GED test form
    If dummy > 0 Then
        calBox.Caption = calBox.Caption + vbNewLine & dummy & " test(s)"
        testDayBool(calBoxNum) = True
    End If
    calBoxNum = calBoxNum + 1
Next
End Sub

Public Function DaysInMonth(month As Integer, year As Integer) As Integer
    firstSerialDate = DateSerial(year, month, 1)
    If month + 1 > 12 Then
        secondSerialDate = DateSerial(year + 1, (month + 1) Mod 12, 1)
    Else
        secondSerialDate = DateSerial(year, month + 1, 1)
    End If
    DaysInMonth = secondSerialDate - firstSerialDate
End Function

Public Function IsOldDate(month As Integer, day As Integer, year As Integer) As Boolean
givenDate = DateSerial(year, month, day)
currentDate = DateSerial(DatePart("yyyy", Now()), DatePart("m", Now()), DatePart("d", Now()))
IsOldDate = givenDate < currentDate
End Function

Public Function IsToday(month As Integer, day As Integer, year As Integer) As Boolean
givenDate = DateSerial(year, month, day)
currentDate = DateSerial(DatePart("yyyy", Now()), DatePart("m", Now()), DatePart("d", Now()))
IsToday = givenDate = currentDate
End Function

Public Function GEDTestDateQ(month As Integer, day As Integer, year As Integer) As Integer
' Returns the number of tests scheduled for a given date
Dim selectedDate As String
selectedDate = "#" & month & "/" & day & "/" & year & "#"
numTests = DLookup("Count(*)", "GEDTestT", "[GEDTestDay] =" & selectedDate)
GEDTestDateQ = numTests
End Function


Public Function getLabelIndex(name As String) As Integer
' Given the name of a label, return its index within variant array of labels on form
getLabelIndex = 0
Dim labelNames As Variant
labelNames = Array("TextA1", "TextA2", "TextA3", "TextA4", "TextA5", "TextA6", "TextA7", _
    "TextB1", "TextB2", "TextB3", "TextB4", "TextB5", "TextB6", "TextB7", _
    "TextC1", "TextC2", "TextC3", "TextC4", "TextC5", "TextC6", "TextC7", _
    "TextD1", "TextD2", "TextD3", "TextD4", "TextD5", "TextD6", "TextD7", _
    "TextE1", "TextE2", "TextE3", "TextE4", "TextE5", "TextE6", "TextE7", _
    "TextF1", "TextF2", "TextF3", "TextF4", "TextF5", "TextF6", "TextF7")

For i = 0 To UBound(labelNames)
    If name = labelNames(i) Then
        getLabelIndex = i
    End If
Next i
End Function

```
