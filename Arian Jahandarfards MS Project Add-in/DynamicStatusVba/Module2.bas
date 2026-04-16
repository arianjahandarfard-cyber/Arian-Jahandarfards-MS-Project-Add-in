Attribute VB_Name = "Module2"
'==============================================================================
' Module: modWorkdayCalc
' Purpose: Working day arithmetic respecting 5-day week + holiday table
' Location: Insert into a standard module in the .xlsm VBA project
'==============================================================================

Option Explicit

Private Const CP_SHEET_NAME As String = "Control Panel"
Private Const HOLIDAY_START_ROW As Long = 10

' Cache the holiday list in memory so we don't re-read the sheet on every call
Private mHolidays() As Date
Private mHolidaysLoaded As Boolean
Private mHolidayCount As Long

Public Sub RefreshHolidayCache()
    ' Call this once before running the propagation engine
    ' Reads all checked (active) holidays from the Control Panel into memory
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim tempDates() As Date
    Dim cnt As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CP_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        mHolidayCount = 0
        mHolidaysLoaded = True
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    If lastRow < HOLIDAY_START_ROW Then
        mHolidayCount = 0
        mHolidaysLoaded = True
        Exit Sub
    End If
    
    ' First pass: count active holidays
    cnt = 0
    For r = HOLIDAY_START_ROW To lastRow
        If ws.Cells(r, 1).Value = True Then
            cnt = cnt + 1
        End If
    Next r
    
    If cnt = 0 Then
        mHolidayCount = 0
        mHolidaysLoaded = True
        Exit Sub
    End If
    
    ReDim tempDates(1 To cnt)
    cnt = 0
    For r = HOLIDAY_START_ROW To lastRow
        If ws.Cells(r, 1).Value = True Then
            cnt = cnt + 1
            ' Use the OBSERVED date (column D), not the calendar date
            tempDates(cnt) = CDate(ws.Cells(r, 4).Value)
        End If
    Next r
    
    mHolidays = tempDates
    mHolidayCount = cnt
    mHolidaysLoaded = True
End Sub

Public Function IsWorkday(dt As Date) As Boolean
    ' Returns True if the date is a working day (Mon-Fri, not a holiday)
    Dim dow As Long
    Dim i As Long
    
    If Not mHolidaysLoaded Then RefreshHolidayCache
    
    dow = Weekday(dt, vbSunday)
    ' Weekend check
    If dow = vbSaturday Or dow = vbSunday Then
        IsWorkday = False
        Exit Function
    End If
    
    ' Holiday check
    If mHolidayCount > 0 Then
        For i = 1 To mHolidayCount
            If Int(dt) = Int(mHolidays(i)) Then
                IsWorkday = False
                Exit Function
            End If
        Next i
    End If
    
    IsWorkday = True
End Function

Public Function AddWorkdays(startDate As Date, workdays As Long) As Date
    ' Adds (or subtracts) a number of working days to a date
    ' If workdays is positive, moves forward; if negative, moves backward
    ' The start date itself is NOT counted as a working day
    Dim result As Date
    Dim direction As Long
    Dim counted As Long
    
    If Not mHolidaysLoaded Then RefreshHolidayCache
    
    result = Int(startDate)  ' Strip time component
    
    If workdays = 0 Then
        AddWorkdays = result
        Exit Function
    End If
    
    direction = IIf(workdays > 0, 1, -1)
    counted = 0
    
    Do While counted < Abs(workdays)
        result = result + direction
        If IsWorkday(result) Then
            counted = counted + 1
        End If
    Loop
    
    AddWorkdays = result
End Function

Public Function CountWorkdays(startDate As Date, endDate As Date) As Long
    ' Counts working days between two dates (exclusive of start, inclusive of end)
    ' This matches how MS Project counts duration
    Dim dt As Date
    Dim cnt As Long
    Dim direction As Long
    
    If Not mHolidaysLoaded Then RefreshHolidayCache
    
    If Int(startDate) = Int(endDate) Then
        CountWorkdays = 0
        Exit Function
    End If
    
    direction = IIf(endDate > startDate, 1, -1)
    cnt = 0
    dt = Int(startDate)
    
    Do
        dt = dt + direction
        If IsWorkday(dt) Then cnt = cnt + 1
    Loop Until Int(dt) = Int(endDate)
    
    CountWorkdays = cnt * direction
End Function

Public Function NextWorkday(dt As Date) As Date
    ' Returns the next working day on or after the given date
    Dim result As Date
    result = Int(dt)
    Do While Not IsWorkday(result)
        result = result + 1
    Loop
    NextWorkday = result
End Function

Public Function PrevWorkday(dt As Date) As Date
    ' Returns the previous working day on or before the given date
    Dim result As Date
    result = Int(dt)
    Do While Not IsWorkday(result)
        result = result - 1
    Loop
    PrevWorkday = result
End Function
