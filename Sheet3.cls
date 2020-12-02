VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
    Application.EnableEvents = False
    Dim bSaveWhenDone As Boolean
    bSaveWhenDone = False
    Dim FormDate As Date
    Dim sFileLocation As String
    If Not Intersect(Target, Range("ShowOrHideSheetTabs")) Is Nothing Then
        Call ShowOrHideSheetTabs
    ElseIf Not Intersect(Target, Range("SetupWorkbook")) Is Nothing Then
        Call SetUpWorkbook
    ElseIf Not Intersect(Target, Range("APIKey")) Is Nothing Then
    'Check the API Key
        GetAPILocation sFileLoc:=sFileLocation
        If Len(sFileLocation) > 0 Then
            Range("APIKey") = sFileLocation
            bSaveWhenDone = True
        Else
            GoTo exit_BDC
        End If
    ElseIf Not Intersect(Target, Range("DataTarget")) Is Nothing Then
    'Check the output file
        GetDataTargetLocation sFileLoc:=sFileLocation
        If Len(sFileLocation) > 0 Then
            Range("DataTarget") = sFileLocation
        Else
            GoTo exit_BDC
        End If
'   MinimumDate (Date) - Restricts the selection of any dates below this date.
'   MaximumDate (Date) - Restricts the selection of any dates above this date.
    ElseIf Not Intersect(Target, Range("StartDate")) Is Nothing Then
    'Check the Start date
        If InStr(1, "Month;Year", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
            'Use the YearMonthForm to get the date as a month/year or year
            FormDate = GetYearMonth("StartDate")
        Else
            'Use the CalendarForm to get the date as a date between the 1st of January 2004 (the starting point of Google Trends data) or two days before the present
            FormDate = CalendarForm.GetDate(SelectedDate:=IIf(IsDate(fvReturnNameValue("StartDate")), fvReturnNameValue("StartDate"), 0) _
                                          , MinimumDate:=#1/1/2004# _
                                          , MaximumDate:=Date - 3 _
                                          , RangeOfYears:=100)
        End If
        If CLng(FormDate) = 0 Then GoTo exit_BDC 'User cancelled the date picker
        Range("StartDate") = FormDate
    ElseIf Not Intersect(Target, Range("EndDate")) Is Nothing Then
    'Check the End date
        If InStr(1, "Month;Year", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
            'Use the YearMonthForm to get the date as a month/year or year
            FormDate = GetYearMonth("EndDate")
            'For the end date, the date must be shifted to the end of the period
            If InStr(1, "Month", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
                FormDate = DateAdd("m", 1, FormDate) - 1
            Else    'i.e., Year
                FormDate = DateSerial(Year(FormDate), 12, 31)
            End If
        Else
            'Use the CalendarForm to get the date as a date between the larger of the 1st of January 2004 or the start date, and two days before the present
            FormDate = CalendarForm.GetDate(SelectedDate:=IIf(IsDate(fvReturnNameValue("StartDate")) And IsDate(fvReturnNameValue("EndDate")) _
                                                            And fvReturnNameValue("EndDate") > fvReturnNameValue("StartDate"), fvReturnNameValue("EndDate"), 0) _
                                          , MinimumDate:=Application.WorksheetFunction.Max(#1/1/2004#, Range("StartDate").Value) _
                                          , MaximumDate:=Date - 2 _
                                          , RangeOfYears:=100)
        End If
        If CLng(FormDate) = 0 Then GoTo exit_BDC 'User cancelled the date picker
        Range("EndDate") = FormDate
    Else
        GoTo exit_BDC
    End If
    If bSaveWhenDone Then DoSave
exit_BDC:
'    Application.EnableEvents = True
    EndGracefully
End Sub
'Not declared private so that it can be called from mFileReadingFunctions.ReadValuesFromFile
Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    
    Dim sMsgBoxString As String 'Used to report date changes to the user
    Dim iDateOffset As Integer  'Creates an offset to adjust a date to the preceding interval (e.g., the preceding month or year)
    
    'Settings for location
    If Not Intersect(Target, Range("Country")) Is Nothing _
        Or Not Intersect(Target, Range("GeographicLevel")) Is Nothing Then
        'Clear Region
        Range("Region").Value = vbNullString
    'Settings for date
    ElseIf Not Intersect(Target, Range("DateResolution")) Is Nothing Then
        If fvReturnNameValue("DateResolution") = "Year" Then
        'User changed date resolution to Year
            Range("StartDate", "EndDate").Style = "InputYear"
            If Not Len(Range("StartDate").Value) = 0 And IsDate(Range("StartDate").Value) Then
                'Adjust the start date to the beginning of the year
                If Not (Range("StartDate").Value) = DateSerial(Year(Range("StartDate").Value), 1, 1) Then
                    Range("StartDate").Value = DateSerial(Year(Range("StartDate").Value), 1, 1)
                    sMsgBoxString = "The Start date has been set to " & Format(Range("StartDate").Value, "dd mmmm yyyy")
                End If
            End If
            If Not Len(Range("EndDate").Value) = 0 And IsDate(Range("EndDate").Value) Then
                'Adjust the end date to the end of the year
                If Not (Range("EndDate").Value) = DateSerial(Year(Range("EndDate").Value), 12, 31) Then
'                    If DateSerial(Year(Range("EndDate").value), 12, 31) > Now Then
'                        iDateOffset = -1
'                    Else
'                        iDateOffset = 0
'                    End If
'                    Range("EndDate").value = DateSerial(Year(Range("EndDate").value) - iDateOffset, 12, 31)
'Application.WorksheetFunction.EoMonth
                    Range("EndDate").Value = fSetMaxDateAfter(d:=Range("EndDate").Value, sDateRes:="Year")
                    sMsgBoxString = sMsgBoxString & IIf(Len(sMsgBoxString) = 0, vbNullString, vbCrLf) _
                     & "The End date has been set to " & Format(Range("EndDate").Value, "dd mmmm yyyy")
                End If
                If Len(sMsgBoxString) > 0 Then _
                MsgBox sMsgBoxString, vbInformation + vbOKOnly, "Dates adjusted for year"
            
            End If
        ElseIf fvReturnNameValue("DateResolution") = "Month" Then
        'User changed date resolution to Month
            Range("StartDate", "EndDate").Style = "InputYearMonth"
            If Not Len(Range("StartDate").Value) = 0 And IsDate(Range("StartDate").Value) Then
                'Adjust the start date to the beginning of the month
                If Not (Range("StartDate").Value) = DateSerial(Year(Range("StartDate").Value), Month(Range("StartDate").Value), 1) Then
                    Range("StartDate").Value = DateSerial(Year(Range("StartDate").Value), Month(Range("StartDate").Value), 1)
                    sMsgBoxString = "The Start date has been set to " & Format(Range("StartDate").Value, "dd mmmm yyyy")
                End If
            End If
            If Not Len(Range("EndDate").Value) = 0 And IsDate(Range("EndDate").Value) Then
                'Adjust the end date to the end of the month
''                If Not (Range("EndDate").value) = DateSerial(Year(Range("EndDate").value), 12, 31) Then
''                    If DateSerial(Year(Range("EndDate").value), 12, 31) > Now Then
''                        iDateOffset = -1
''                    Else
''                        iDateOffset = 0
''                    End If
''                    Range("EndDate").value = DateSerial(Year(Range("EndDate").value) - iDateOffset, 12, 31)
''                End If
                If Not (Range("EndDate").Value) = Application.WorksheetFunction.EoMonth(fvReturnNameValue("EndDate"), 0) Then
                    Range("EndDate").Value = fSetMaxDateAfter(d:=Range("EndDate").Value, sDateRes:="Month")
                    sMsgBoxString = sMsgBoxString & IIf(Len(sMsgBoxString) = 0, vbNullString, vbCrLf) _
                     & "The End date has been set to " & Format(Range("EndDate").Value, "dd mmmm yyyy")
                End If
            End If
            If Len(sMsgBoxString) > 0 Then _
            MsgBox sMsgBoxString, vbInformation + vbOKOnly, "Dates adjusted for month"
        
        ElseIf InStr(1, "Day;Week", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
        'User changed date resolution to Week or Day
            Range("StartDate", "EndDate").Style = "InputFullDate"
'            SetDataValidationStartDW
'            SetDataValidationEndDW
            If fvReturnNameValue("DateResolution") = "Week" Then
                If Day(Range("StartDate").Value) <> 1 Then
                    Range("StartDate").Value = Range("StartDate").Value - Day(Range("StartDate").Value) + 1
                    sMsgBoxString = "The Start date has been set to " & Format(Range("StartDate").Value, "dd mmmm yyyy")
                End If
                If Day(Range("EndDate").Value) <> 7 Then
                    Range("EndDate").Value = fSetMaxDateAfter(d:=Range("EndDate").Value, sDateRes:="Week")
                    sMsgBoxString = "The End date has been set to " & Format(Range("EndDate").Value, "dd mmmm yyyy")
                End If
                If Len(sMsgBoxString) > 0 Then _
                MsgBox sMsgBoxString, vbInformation + vbOKOnly, "Dates adjusted for week"
            End If
        End If
    'Settings for Start- and end date
    ElseIf Not Intersect(Target, Range("StartDate")) Is Nothing _
        Or Not Intersect(Target, Range("EndDate")) Is Nothing Then
        'Process incomplete dates when the resolution is set to year or month
        'i.e., if the calendar dialog is used to set the date as a month or year, then convert it to a full date
        With Target
            If fvReturnNameValue("DateResolution") = "Year" Then
                If IsNumeric(.Value2) And .Value2 >= 2004 And .Value2 <= Year(Now) Then
                    If Not Intersect(Target, Range("StartDate")) Is Nothing Then
                        .Value2 = DateSerial(.Value2, 1, 1)
'''                        SetDataValidationEndY   'Reset the data validation for the end year
                    ElseIf Not Intersect(Target, Range("EndDate")) Is Nothing Then
                        .Value2 = DateSerial(.Value2, 12, 31)
                    End If
                End If
            ElseIf fvReturnNameValue("DateResolution") = "Month" Then
                On Error Resume Next
                Dim V As Variant
                V = .Value
                .Value = DateValue(.Value)
                If Err.Number <> 0 Then
                    Err.Clear
                    Target.Value = V
                End If
                On Error GoTo 0
            End If
        End With
'    ElseIf Not Intersect(Target, Range("DataTarget")) Is Nothing Then
'        FileDirCheck file, fvReturnNameValue("DataTarget")
    End If
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Turn off the can't paste message so that it doesn't overstay its purpose
    'For what it's worth, this does not wipe out the undo stack
    With Application
        If .StatusBar = sPasteDisallowMsg Then .StatusBar = vbNullString
    End With
End Sub
