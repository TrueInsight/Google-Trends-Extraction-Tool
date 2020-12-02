VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    If bProductionVersion Then
        'If the user navigated to the Google Trends Web sheet manually, then hide the category sheets
        If Sheet12.Visible <> xlSheetVeryHidden Then Sheet12.Visible = xlSheetVeryHidden
        If Sheet8.Visible <> xlSheetVeryHidden Then Sheet8.Visible = xlSheetVeryHidden
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
    Application.EnableEvents = False
    Dim FormDate As Date
    Dim sFileLocation As String

    If Not Intersect(Target, Range("WCategory")) Is Nothing Then
        ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
        With Sheet8
            .Visible = xlSheetVisible
            .Activate
        End With
        'GoTo exit_WBDC
    ElseIf Not Intersect(Target, Range("WDataTarget")) Is Nothing Then
    'Check the output file
        GetDataTargetLocation sFileLoc:=sFileLocation
        If Len(sFileLocation) > 0 Then
            Range("WDataTarget") = sFileLocation
        Else
            GoTo exit_WBDC
        End If
    ElseIf Not Intersect(Target, Range("WStartDate")) Is Nothing Then
    'Check the Start date
'        If InStr(1, "Month;Year", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
'            'Use the YearMonthForm to get the date as a month/year or year
'            FormDate = GetYearMonth("WStartDate", "W")
'        Else
            'Use the CalendarForm to get the date as a date between the 1st of January 2004 (the starting point of Google Trends data) or two days before the present
            FormDate = CalendarForm.GetDate(SelectedDate:=IIf(IsDate(fvReturnNameValue("WStartDate")), fvReturnNameValue("WStartDate"), 0) _
                                          , MinimumDate:=#1/1/2004# _
                                          , MaximumDate:=Date - 2 _
                                          , RangeOfYears:=100)
'        End If
        If CLng(FormDate) = 0 Then GoTo exit_WBDC 'User cancelled the date picker
        Range("WStartDate") = FormDate
    ElseIf Not Intersect(Target, Range("WEndDate")) Is Nothing Then
    'Check the End date
'        If InStr(1, "Month;Year", fvReturnNameValue("DateResolution"), vbTextCompare) > 0 Then
'            'Use the YearMonthForm to get the date as a month/year or year
'            FormDate = GetYearMonth("WEndDate", "W")
'        Else
'            'Use the CalendarForm to get the date as a date between the larger of the 1st of January 2004 or the start date, and two days before the present
            FormDate = CalendarForm.GetDate(SelectedDate:=IIf(IsDate(fvReturnNameValue("WStartDate")) And IsDate(fvReturnNameValue("WEndDate")) _
                                                            And fvReturnNameValue("WEndDate") > fvReturnNameValue("WStartDate"), fvReturnNameValue("WEndDate"), 0) _
                                          , MinimumDate:=Application.WorksheetFunction.Max(#1/1/2004#, Range("WStartDate").Value) _
                                          , MaximumDate:=Date - 2 _
                                          , RangeOfYears:=100)
'        End If
        If CLng(FormDate) = 0 Then GoTo exit_WBDC 'User cancelled the date picker
        Range("WEndDate") = FormDate
    'Settings for location
    Else
        GoTo exit_WBDC
    End If

exit_WBDC:
'    Application.EnableEvents = True
    EndGracefully
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    Application.EnableEvents = False
    'Settings for location
    If Not Intersect(Target, Range("WCountry")) Is Nothing _
        Or Not Intersect(Target, Range("WGeographicLevel")) Is Nothing Then
        'Clear Region
        Range("WRegion").Value = vbNullString
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


