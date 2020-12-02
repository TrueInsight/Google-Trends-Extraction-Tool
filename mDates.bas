Attribute VB_Name = "mDates"
'Google Trends Extended for Health Information Extraction Tool
'e-mail: jacques.raubenheimer@ sydney.edu.au
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
Option Explicit
Option Private Module
Option Base 1
Private Enum PeriodStartOrEnd
    PeriodStart = 1
    PeriodEnd = 2
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The functions in this module allow for the building of arrays of dates  '
' to be used in the query specifications and data extraction             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetYearMonth(ByRef sRangeName As String _
           , Optional ByVal sWorksheetPrefix As String = vbNullString) As Date
'This function reads the existing Start or End date from the specified range,
' and then launches the YearMonthForm to allow the user to select a year/month,
' which is then written back to the rance

    Dim MinYear As Integer
    Dim curYear As Integer
    
    If Not IsDate(fvReturnNameValue(sRangeName)) Or sRangeName = sWorksheetPrefix & "StartDate" Then
            MinYear = 2004
    ElseIf sRangeName = sWorksheetPrefix & "EndDate" Then
        MinYear = Year(fvReturnNameValue(sWorksheetPrefix & "StartDate"))
    End If
    If IsDate(fvReturnNameValue(sRangeName)) Then
        curYear = Year(fvReturnNameValue(sRangeName))
    Else
        curYear = MinYear
    End If
    
    Dim fmYM As YearMonthForm
    Set fmYM = New YearMonthForm
    With fmYM
        .CallingCell = sRangeName
        .MinYear = MinYear
        .StartingYear = curYear '.StartingYear = minYear
        .StartingMonth = IIf(IsDate(fvReturnNameValue(sRangeName)) _
                            , Format(fvReturnNameValue(sRangeName), "MMMM"), "January")
        .YearAndMonth = (fvReturnNameValue("DateResolution") = "Month" Or sWorksheetPrefix = "W")
        .Show
'        Debug.Print .SelectedDate
        If .FormCancel Then
            GetYearMonth = 0
        Else
            GetYearMonth = .SelectedDate
        End If
        Unload fmYM
        'if
    End With
End Function

Public Function ReturnArrayOfMonths() As Variant
'This function creates a 12 x 1 array containing each month of the year
'This is used for the YearMonthForm

    Dim sMonths() As Variant
    ReDim sMonths(1 To 12)
    sMonths(1) = "January"
    sMonths(2) = "February"
    sMonths(3) = "March"
    sMonths(4) = "April"
    sMonths(5) = "May"
    sMonths(6) = "June"
    sMonths(7) = "July"
    sMonths(8) = "August"
    sMonths(9) = "September"
    sMonths(10) = "October"
    sMonths(11) = "November"
    sMonths(12) = "December"
    ReturnArrayOfMonths = sMonths
End Function

Public Function ReturnYearsForGoogleTrends(Optional ByRef bUseFirstYear As Boolean) As Variant
'This function builds an array of the actual years included in the date range specification,
' from the year of the Start date, to the year of the end date
'This is used for the YearMonthForm

    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Function
    
    Dim iYearInterval As Integer    'The number of years covered by the start and end dates
    Dim iCurrentYear As Integer     'The year of the end date
    Dim iFirstYear As Integer       'The year of the start date
    Dim arYears() As Variant        'A temporary array created to list each year
    Dim i As Integer                'temporary counter
    
    If IsDate(fvReturnNameValue("StartDate")) And Not bUseFirstYear Then
        iFirstYear = Year(fvReturnNameValue("StartDate"))
    Else
        iFirstYear = 2004
    End If
    
    iCurrentYear = Year(Now)
    iYearInterval = iCurrentYear - iFirstYear + 1
    
    ReDim arYears(1 To iYearInterval)
    For i = 1 To iYearInterval
        arYears(i) = iFirstYear + i - 1
    Next i
    
    ReturnYearsForGoogleTrends = arYears

End Function
Private Sub TestDP()
'Just a short procedure to show how the DatePart function operates. Not used
Debug.Print DatePart("yyyy", Date)  'Year
Debug.Print DatePart("q", Date)     'Quarter
Debug.Print DatePart("m", Date)     'Month
Debug.Print DatePart("y", Date)     'Day of year
Debug.Print DatePart("d", Date)     'Day, i.e., day of month
Debug.Print DatePart("w", Date)     'Weekday, start with 1 for Sunday if FirstDayOfWeek is not specified otherwise
Debug.Print DatePart("ww", Date)    'Week [of year], starting with week in which Jan 1 falls when FirstWeekOfYear is not specified otherwise

Debug.Print DateAdd("ww", 1, Date)
End Sub

Public Function ReturnPeriodsAsArray(Optional bEndOnError As Boolean = False) As Variant
'This function builds a two-dimensional array with a start date and end date for every period.
'This is used for the Google Extended for Health queries, but not the Google Trends Web queries.
'The Start and End date as specified in the query specification are used,
' and then the date resolution is used to divide the total time up into distinct periods.
'This is the core work that reproduces the multiple sampling strategy I devised to get around the Google Trends data caching problem,
' and which I hope to publish in a separate publication.

    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Function
    
    Dim dStart As Date
    Dim dEnd As Date
    Dim sFreq As String
    Dim iNper As Integer
    Dim aPeriods() As Variant
    Dim i As Integer
    Dim sDateInterval As String
    Dim dAbsoluteStart As Date
    Dim dAbsY As Integer
    Dim dAbsM As Integer
    Dim dAbsD As Integer
    
    dStart = fvReturnNameValue("StartDate")
    If Not IsDate(dStart) Or dStart < #1/1/2004# Then
    'The start date specified on the 'Query selection' sheet and read into dStart is not correct
        If Not IsDate(dStart) Then
            MsgBox "The value of '" & dStart & "' for the starting date is not a valid date.", vbCritical + vbOKOnly, "Invalid Start date"
        ElseIf dStart < #1/1/2004# Then
            MsgBox "The value of '" & dStart & "' for the starting date is before the 1st of January, 2004. No Google Trends data exist before this date.", vbCritical + vbOKOnly, "Start date too early"
        End If
        If bEndOnError Then EndGracefully
    End If
    
    dEnd = fvReturnNameValue("EndDate")
    If Not IsDate(dEnd) Or dEnd < dStart Or dEnd > Date - 2 Then
    'The End date specified on the 'Query selection' sheet and read into dEnd is not correct
        If Not IsDate(dEnd) Then
            MsgBox "The value of '" & dEnd & "' for the Ending date is not a valid date.", vbCritical + vbOKOnly, "Invalid End date"
        ElseIf dEnd > Date - 2 Then
            MsgBox "The value of '" & dEnd & "' for the Ending date is after two days before the current date. Google Trends Extended data are not available for this date.", vbCritical + vbOKOnly, "End date too late"
        End If
        If bEndOnError Then EndGracefully
    End If
    
    iNper = fvReturnNameValue("Periods")
    If Not iNper > 1 Then
    'The number of periods specified on the 'Query selection' sheet and read into iNPer is not correct
        MsgBox "The value of '" & iNper & "' for the number of periods is not correct.", vbCritical + vbOKOnly, "Invalid period specification"
        If bEndOnError Then EndGracefully
    End If
    
    sFreq = fvReturnNameValue("DateResolution")
    If Not InStr(1, "Day;Week;Month;Year", sFreq, vbTextCompare) > 0 Then
    'The Frequency specified on the 'Query selection' sheet and read into sFreq is not correct
        MsgBox "The frequency value of '" & sFreq & "' for the period is not correct.", vbCritical + vbOKOnly, "Invalid period resolution"
        If bEndOnError Then EndGracefully
    End If
    
    'To determine the correct bounds for the date periods, the type of time "unit" must be determined
    'Simply use this first step to return a date formatting string that Excel recognises (m and d are already correct, y and w require a little tweaking)
    sDateInterval = LCase(Left$(sFreq, 1))
    'Adjust sDateInterval for the year/week
    If sDateInterval = "y" Then
        sDateInterval = "yyyy"
    ElseIf sDateInterval = "w" Then
        sDateInterval = "ww"
    End If
    
    'Set a fixed point in time from which all the data calculations will be done
    'This is stored in dAbsoluteStart
    Select Case sDateInterval
    Case "yyyy"
        dAbsY = DatePart(sDateInterval, dStart)
        dAbsM = 1
        dAbsD = 1
    Case "m"
        dAbsY = DatePart("yyyy", dStart)
        dAbsM = DatePart(sDateInterval, dStart)
        dAbsD = 1
    Case "ww"
        Dim tmpDay As Integer
        tmpDay = DatePart("w", dStart) - 1          'Use this to find the start of the week in which the date occurs, which may be in the preceding year and/or month
        dAbsY = DatePart("yyyy", dStart - tmpDay)
        dAbsM = DatePart("m", dStart - tmpDay)
        dAbsD = DatePart("d", dStart - tmpDay)
    Case "d"
        dAbsY = DatePart("yyyy", dStart)
        dAbsM = DatePart("m", dStart)
        dAbsD = DatePart(sDateInterval, dStart)
    End Select
    dAbsoluteStart = DateSerial(dAbsY, dAbsM, dAbsD)
    
    'The first dimension of aPeriods only has two values--1 for the start and 2 for the end
    'The second dimension is as long as the number of periods specified on the query sheet
    ' i.e., this returns a start and end date for each period
    ReDim aPeriods(PeriodStart To PeriodEnd, 1 To iNper)
    
    'Set the start of the first period
    aPeriods(PeriodStart, 1) = Format(dStart, "yyyy-mm-dd")
    'Set the end of the last period
    aPeriods(PeriodEnd, iNper) = Format(dEnd, "yyyy-mm-dd")
    'The first period encompasses the full time range
    
    'Assign the Start and/or End of each period to the array element
    For i = 1 To iNper
        If i = 1 Then
            'Only set the end of the first period
            aPeriods(PeriodEnd, 1) = Format(DateAdd(sDateInterval, 1, dAbsoluteStart) - 1, "yyyy-mm-dd")
        ElseIf i = iNper Then
            'Only set the start of the last period
            aPeriods(PeriodStart, iNper) = Format(DateAdd(sDateInterval, iNper - 1, dAbsoluteStart), "yyyy-mm-dd")
        Else
            'Set the start and end of all intervening periods
            aPeriods(PeriodStart, i) = Format(DateAdd(sDateInterval, i - 1, dAbsoluteStart), "yyyy-mm-dd")
            aPeriods(PeriodEnd, i) = Format(DateAdd(sDateInterval, i, dAbsoluteStart) - 1, "yyyy-mm-dd")
        End If
    Next i
    
    ReturnPeriodsAsArray = aPeriods
    
End Function
