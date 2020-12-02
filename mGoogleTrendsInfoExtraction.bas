Attribute VB_Name = "mGoogleTrendsInfoExtraction"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the code which does the actual extraction and parsing      '
' of data from the Google Trends Extended for Health (a.k.a. Google Flu) service. '
' It calls procedures from several of the other modules.                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub DrawSample()
'This is the main procedure that is launched when the user clicks the Extract Data button on the
' Google Trends Extended Health worksheet (Sheet3).
'It first does a complete check on each component needed for the extraction string as specified on the input worksheet,
' and then launches the procedure that sends the request to Google and builds the workbook with the returned data.
    
    StartSmoothly
    
    'Read the quotas set by Google for error checking. These are read from the worksheet.
    SetHardBounds
    
    'Error checking on every component of the specification
    If Not fDoAllErrorChecking Then EndGracefully
    
    'If the error checking passes all clear, then build the query request
    'A one-dimensional array is returned where each element of the array is a complete URL to be passed to Google
    Dim vQueryArray() As Variant
    vQueryArray = fBuildRequestArray
        
    'Now that all the queries are stored in a single array variable, pass each one to Google and parse the results
    
    'Check whether the extracted data must be parcelled out to separate worksheets, and if so, on what basis
    Dim iMultiSheet As BuildSheets
    If fvReturnNameValue("IsMultiRegionRequest") And fvReturnNameValue("IsMultiTermRequest") Then
        'Although I originally envisioned coding the tool to do this,
        ' I have not permitted it at present, as the restrictions on the data returned
        ' by the API (primarily, the limit of 2000 data points in the JSON string)
        ' make it very impractical
        iMultiSheet = BuildSheetsByBoth
    ElseIf fvReturnNameValue("IsMultiRegionRequest") Then
        iMultiSheet = BuildSheetsByRegion
    ElseIf fvReturnNameValue("IsMultiTermRequest") Then
        iMultiSheet = BuildSheetsByQueryTerm
    Else
        iMultiSheet = BuildSheetsNone
    End If
    
    'Invoke the workhorse procedure that builds the target workbook, and pulls each URL's data into the relevant worksheets
    DrawSampleDataFromGoogle vURLArray:=vQueryArray, iDoMultiple:=iMultiSheet
    
    'Technically, this point should not be reached in a normal run, as EndGracefully is called at the end of CompleteReporting
    If bShowCompletionMsgBoxes Then EndGracefully
    
End Sub

'Sub DrawSampleDataFromGoogle(ByRef vURLArray() As Variant,Optional ByVal sSubstance As String = vbNullString, Optional ByVal tf As TimeFrame = TimeFrameDay)
Sub DrawSampleDataFromGoogle(ByRef vURLArray() As Variant _
                           , Optional ByRef iDoMultiple As BuildSheets = BuildSheetsNone)
'This is the workhorse procedure that takes the array of URLs which has been created, feeds them to the API
' and then parses back the results. It calls several procedures to accomplish this.
    
    Dim i As Integer                            'General purpose counter
    Dim j As Integer                            'General purpose counter
    
    Dim sURL As String                          'String to which each URL will be written before passing to the fGetData function
    Dim iQueryCounter As Integer                'Counts how many requests are done so that the limit of 200 queries per second is not exceeded
    Dim iTotNumSamplingsRequested As Integer    'Calculates the total number of samplings needed to obtain the requested number of samples
    Dim iQTCount As Integer                     'Counts the number of query terms, so that a corresponding number of worksheets can be created
    Dim iRegionCount As Integer                 'Counts the number of regions for which queries are submitted
    Dim wbkNew As Workbook                      'Stores the data extracted from Google Trends
    Dim wksGTData() As Worksheet                'Worksheet(s) in wbkNew where the actual data are written
    Dim wksSummary() As Worksheet               'Worksheet(s) in wbkNew where the data from wksGTData are summarised: Mean per time point, as well as N/median/min/max/range/SD
    Dim wksMainSummary As Worksheet             'Worksheet in wbkNew that combines all the average series from the wksSummary array
                                                'Note that if only one term is requested for one region, then wksMainSummary=wksSummary(1)
    Dim wksSamplingSummary As Worksheet         'Worksheet which summarises sampling adequacy for the various terms [Added in version 2.0.0]
    Dim wksQuerySpec As Worksheet               'Worksheet in wbkNew that stores the complete query specifications for auditing purposes
    Dim vQueryList() As Variant                 'Stores the list of queries for use in identifying each worksheet
    
    Dim sResult As String                       'Captures the data returned by the Google Trends server in the HTTP request
    Dim V() As Variant                          'Variant array used to extract dates and values from the sResult string, and then transpose-dumped into the wksGTData worksheet
    Dim vTmp() As Variant                       'v is returned as a two-dimensional array
                                                ' (when only one geographical region and only one query term is specified,
                                                '  then the first column still contains the dates, and the second column, the values.
                                                '  When more than one region/query term are specified, then each column (from 2 onwards)
                                                '  contains values for one region/term
                                                'vTmp then allows one column to be extracted from v to be written to one worksheet at a time
    Dim vTransposeResult() As Variant           'Created this 2020-09 to use in the TransposeArray output
    Dim rng As Range                            'range variable used for afore-mentioned transpose-dump
    
    Dim lDateCounter As Long                    'These two variables are used to create a list of dates from the start to the end date,
    Dim lDateEnd As Long                        ' with the interval as defined in DateResolution
    
    Dim iMaxRow As Integer                      'Sets the number of rows of extracted data for writing the formulas, etc.
    Dim sTimeInterval As String                 'writes the time interval of the query (day/week/month/year)
    Dim tf As TimeFrame                         'Reads the time frame that the user has chosen
    
    Dim StartDate As Date                       'The date retrieved from the JSON string
    
    Dim TimeProcessStart As Single              'Stores the system time to measure how long the extraction process takes
    Dim TimePerQuery As Single                  'Stores the system time to check that the per-second quota is not exceeded
    Dim iQPS As Integer                         'Counts the queries to be checked against the above time
    Dim sMultipleTitle As String                'Used to label sheets when data are parcelled out to separate worksheets
    Dim sMultipleColHeads() As String           'Used to label columns when data are parcelled out to separate worksheets
    Dim iMultipleCount As Integer               'Notes how many additional sheets are required
    Dim iMultiCounter As Integer                'Counts through the additional sheets for the data parcelling
    Dim iDataColumnCounter As Integer           'Keeps track of which column the data should be written to
    
'    wksCompleteSamples is not used in the main application version, but was used in testing. It remains here (commented out) for possible documentation use.
'    Dim wksCompleteSamples As Worksheet        'Worksheet in wbkNew where the individual "pairs" of samplings from wksGTData are collapsed into complete samples (in Procedure CreateCombinedWorksheet)
    
    TimeProcessStart = Timer
    
    'Determine the time frame (the string manipulation method below is simpler than the Select Case method,
    'But TestSpeedOfTimeFrameMethods shows that the string manipulation method is considerably slower, so I am leaving it out
    'tf = InStr(1, Left$(fvReturnNameValue("DateResolution"), 1), "DWMY", vbTextCompare)
    Select Case fvReturnNameValue("DateResolution")
    Case "Day"
        tf = TimeFrameDay
        sTimeInterval = "d"
    Case "Week"
        tf = TimeFrameWeek
        sTimeInterval = "ww"
    Case "Month"
        tf = TimeFrameMonth
        sTimeInterval = "m"
    Case "Year"
        tf = TimeFrameYear
        sTimeInterval = "yyyy"
    End Select
    
''    'Determine the number of queries we will be sending
''    Dim rr As Long
''    rr = UBound(vURLArray) - LBound(vURLArray) + 1

    'Determine the number of query terms requested
    iQTCount = Application.WorksheetFunction.CountA(Range("SearchTermList"))
    'Determine the number of regions requested (if region comparison is requested)
    iRegionCount = Application.WorksheetFunction.CountIf(Range("ISO3166ParentSubdivisions"), fvReturnNameValue("Descriptor"))
    
    'Get the total number of samples requested
    iTotNumSamplingsRequested = fvReturnNameValue("Samples") * 2 - 1
        
    'First collect all the query terms into the array vQueryList.
    'This is different to BuildQueryList which returns a single string with all the query terms preceded by "&term="
    'Check through the whole range, just to be sure
    'Trim away leading and trailing spaces as well
    For i = 1 To 30
        If Len(Trim$(fvReturnNameValue(sName:="SearchTerm" & Format(i, "00"), bCheckForActiveWorkbook:=False))) > 0 Then
            ReDim Preserve vQueryList(1 To i)
            vQueryList(i) = Trim$(fvReturnNameValue(sName:="SearchTerm" & Format(i, "00"), bCheckForActiveWorkbook:=False))
        End If
    Next i
    
    If UBound(vQueryList) <> iQTCount Then
    'There is an error, because the count of nonempty cells is not the same as the number of elements in the query terms array
        MsgBox "The list of query terms could not be successfully reconciled." _
        & vbCrLf & "Check that there are no cells containing only spaces in the list of query terms, " _
        & "or for other possible errors." _
        , vbCritical + vbOKOnly, "Query terms not loaded to array"
        EndGracefully
    End If
    
    'Create a summary of the query and data extraction for auditing purposes
    'If a GeographicLevel of "Region" is specified, then store the region or regions in an array,
    ' otherwise just set the array to a null value
    Dim vRegions() As Variant
    If fvReturnNameValue("GeographicLevel") <> "Region" Then
        ReDim vRegions(1 To 1)
        vRegions(1) = vbNullString
    Else
        If iDoMultiple = BuildSheetsByRegion Or iDoMultiple = BuildSheetsByBoth Then
            vRegions = Range("GeoAllRegionsForCountry")
        Else
            ReDim vRegions(1 To 1)
            vRegions(1) = fvReturnNameValue("Region")
        End If
    End If
    
    'Get the Start and End dates
    Sheet3.Activate
    lDateCounter = fvReturnNameValue(sName:="StartDate", bCheckForActiveWorkbook:=False)
    lDateEnd = fvReturnNameValue(sName:="EndDate", bCheckForActiveWorkbook:=False)
    
    'Build appropriate header sets for multiple-component queries
    Select Case iDoMultiple
    Case BuildSheetsNone
        sMultipleTitle = vbNullString
        iMultipleCount = 1
        ReDim sMultipleColHeads(1 To 1)
        sMultipleColHeads(1) = fvReturnNameValue("SearchTerm01")
    Case BuildSheetsByRegion
        sMultipleTitle = "Region"
        iMultipleCount = iRegionCount
        ReDim sMultipleColHeads(1 To iMultipleCount)
        For i = LBound(vRegions) To UBound(vRegions)
            sMultipleColHeads(i) = vRegions(i, 1)
        Next i
    Case BuildSheetsByQueryTerm
        sMultipleTitle = "Query Term"
        iMultipleCount = iQTCount
        ReDim sMultipleColHeads(1 To iMultipleCount)
        For i = LBound(vQueryList) To UBound(vQueryList)
            sMultipleColHeads(i) = vQueryList(i)
        Next i
'    Case BuildSheetsByBoth
        'Do Nothing
        'At the moment, the on-sheet error checking will not allow a multi-term, multi-region search.
        'If there are x regions and y query terms,
        ' then it would generate have x*y sampling sheets, with x*y summary sheets,
        ' as well as x summary sheets showing all terms for each region,
        ' and y symmary sheets showing all regions for each term.
        ' This will also generate a massive number of sheets (N=x*y*2 + x + y + 1).
        ' Furthermore, the restriction of 2000 data points returned per query makes this impractical to implement.
    End Select
    
    'Add the new workbook to which the extracted data is going to be written
    Set wbkNew = Application.Workbooks.Add
    With wbkNew
        'Add a worksheet that will bring together the means from all the worksheet summaries
        Set wksMainSummary = .Sheets(1)
        
        'Add the sampling adequacy summary sheet
        Set wksSamplingSummary = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        wksSamplingSummary.Name = "Sampling summary"
        
        'Add the query specification sheet for auditing purposes
        Set wksQuerySpec = .Sheets.Add(before:=.Sheets(1))
        
        'Prepare for adding the sheets which will store and summarize the returned data
        ReDim wksSummary(1 To iMultipleCount)
        ReDim wksGTData(1 To iMultipleCount)
        
        'If there is only one query term, then a number of separate worksheets are not needed
        If iDoMultiple = BuildSheetsNone Then
            Set wksSummary(1) = wksMainSummary
            wksSummary(1).Name = "Summary"
            Set wksGTData(1) = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            wksGTData(1).Name = "Samplings"
        Else
            wksMainSummary.Name = "Means of all " & sMultipleTitle & " series"
            For i = 1 To iMultipleCount
                'Add a worksheet that will contain a calculated summary of all the extracted data
                Set wksSummary(i) = .Sheets.Add(After:=.Sheets(.Sheets.Count))
                wksSummary(i).Name = fReturnSafeWorksheetName(sStartName:="Summary-" & sMultipleColHeads(i), sExistingName:=wksSummary(i).Name, wb:=wbkNew)
                'Add a worksheet to which each extracted sampling will be written (the raw data)
                Set wksGTData(i) = .Sheets.Add(After:=.Sheets(.Sheets.Count))
                wksGTData(i).Name = fReturnSafeWorksheetName(sStartName:="Samplings-" & sMultipleColHeads(i), sExistingName:=wksSummary(i).Name, wb:=wbkNew)
            Next i
        End If
    End With    'wbkNew
    
    'Write a summary of the query specification to a worksheet for later auditing purposes
    SetUpQuerySpecSheet wksQuerySpec:=wksQuerySpec _
                       , vQueryList:=vQueryList _
                       , vRegions:=vRegions _
                       , vURLs:=vURLArray
    
    'Prepare the basic layout of the Sampling summary sheet
    SetUpSamplingSummarySheet wksSS:=wksSamplingSummary _
                            , iMultipleCount:=iMultipleCount _
                            , wksSmy:=wksSummary _
                            , sColHeads:=sMultipleColHeads
    
    'Now write the dates to the left column of the Sampling worksheet
    'It is written to wksMainSummary only, and then copied across to all others below.
    j = 1
    wksMainSummary.Cells(1, 1) = "Date"
    Do While lDateCounter <= lDateEnd
        j = j + 1
        wksMainSummary.Cells(j, 1).Value = lDateCounter
        lDateCounter = DateAdd(sTimeInterval, 1, lDateCounter)
    Loop
    iMaxRow = j
    
    'Format the above-mentioned dates appropriately
    'Reuse sTimeInterval here to set the date formatting
    If tf = TimeFrameDay Or tf = TimeFrameWeek Then
        sTimeInterval = "yyyy-mm-dd"
    ElseIf tf = TimeFrameMonth Then
        sTimeInterval = "yyyy-mm"
    ElseIf tf = TimeFrameYear Then
        sTimeInterval = "yyyy"
    End If
    
    wksMainSummary.Range(wksMainSummary.Cells(2, 1), wksMainSummary.Cells(iMaxRow, 1)).NumberFormat = sTimeInterval
    
    'Copy the dates to the various worksheets which will contain the data
    ' (this is faster than repeating the whole process above for each sheet)
    Set rng = wksMainSummary.Range(wksMainSummary.Cells(1, 1), wksMainSummary.Cells(iMaxRow, 1))
    If iDoMultiple = BuildSheetsNone Then
        rng.Copy Destination:=wksGTData(1).Cells(1, 1)
    Else
        For i = 1 To iMultipleCount
            rng.Copy Destination:=wksGTData(i).Cells(1, 1)
            rng.Copy Destination:=wksSummary(i).Cells(1, 1)
        Next i
    End If
    Set rng = Nothing
    
    'Save the workbook once before we start writing the data
'    wbkNew.SaveAs fvReturnNameValue(sName:="DataTarget", bCheckForActiveWorkbook:=False)
    SaveWithErrorHandling iSaveOrSaveAs:=2 _
                        , wbk:=wbkNew _
                        , sFilePath:=fvReturnNameValue(sName:="DataTarget", bCheckForActiveWorkbook:=False) _
                        , bAddToMRU:=False
    
    iMultiCounter = 1       'Set to parcel data out for the first range (if specified)
    iDataColumnCounter = 0
    
    'Retrieve all the samplings and write them to the right worksheets
    For j = LBound(vURLArray) To UBound(vURLArray)
        
        Application.StatusBar = "Fetching sampling " & j & " (of " _
            & iTotNumSamplingsRequested * (iMultipleCount / iQTCount) _
            & ") from Google Trends" _
            & IIf(iDoMultiple = BuildSheetsByRegion _
            Or iDoMultiple = BuildSheetsByBoth, " (region " & iMultiCounter & ")", "")
        
        'Extract the data for URL(j)
        TimePerQuery = Now
        On Error Resume Next
        'StartDate is set within this function (it is passed ByRef) because the starting date will differ for individual samplings
        V = fvGetAndParseGoogleData(sURL:=CStr(vURLArray(j)), sDate:=StartDate, iNTerms:=iQTCount)
        If Err.Number <> 0 Then
            'If an error occurred that was not dealt with, then stop processing
            Err.Clear
            Exit For
        ElseIf bQuotaExceeded Or bOtherHTTPError Or bInvalidArgumentError Then
            'If the quota exceeded error is returned, or an unrecoverable HTTP error is encountered, or an invalid query is submitted [added 2020-09-30]
            ' then there is no data to parse, so exit the loop (stops requesting more samples)
            ' This causes CompleteReporting to be invoked, so that the existing reporting file can be tidied up
            Exit For
        End If
        On Error GoTo 0
    
        'Find the row on which to start putting in the data.
        ' The samplings do not all start on the same row.
        ' They will, however, either start on the first row, or end on the last row.
        'Because StartDate is passed to sDate byRef, this also sets StartDate
        ' which is then used below to find the right row in which to put the data.
        
        'First increment the data column counter
        iDataColumnCounter = iDataColumnCounter + 1
        
        'If the query specification requests a single location and a single term,
        '  then the array v contains data for that sampling (i.e., first dimension contains two elements)
        '  Each extraction of a URL (counted by j) is written as a sampling to the (only) Samplings sheet
        'If the query specification is for one term and multiple locations,
        '  then the array v contains data for that sampling and that particular location (i.e., first dimension contains two elements)
        '  Each extraction of a URL (counted by j) is written to the next column in that item in the series of Samplings sheets
        '  (i.e., the URLs contain all samplings for region 1, then all samplings for region 2, etc.)
        'If the query specification is for multiple terms,
        '  then the array v contains data for all terms, for that sampling and that particular location (i.e., first dimension contains elements equal to the number of terms +1)
        '  Each extraction of a URL (counted by j) is written, piecemeal, to all of the Samplings sheets
        'As indicated above, the situation where there are multiple query terms as well as multiple regions is not allowed
        If iDoMultiple = BuildSheetsNone Or iDoMultiple = BuildSheetsByRegion Then
            
            'Add a column head for this data extraction for each data worksheet
            wksGTData(iMultiCounter).Cells(1, iDataColumnCounter + 1).Value = "Sampling " & iDataColumnCounter
            lDateCounter = Application.match(Application.EDate(StartDate, 0) _
                                           , wksGTData(iMultiCounter).Range(wksGTData(iMultiCounter).Cells(1, 1) _
                                           , wksGTData(iMultiCounter).Cells(iMaxRow, 1)), 0)
            Set rng = wksGTData(iMultiCounter).Cells(lDateCounter, iDataColumnCounter + 1)
            
            'Extract the right column from v so that it can be written to the worksheet
            Call TransferOneDimensionFromArray(vSourceArr:=V _
                                                 , vDestinationArr:=vTmp _
                                                 , iDimOneOrDimTwo:=2 _
                                                 , lWhichItem:=2 _
                                                 , lStartPoint:=2 _
                                                 , lEndPoint:=UBound(V, 2) _
                                                 , lDestArrBase:=1)
            
            'Write the data to the worksheet
            'rng.Resize(UBound(vTmp), 1).Value = Application.Transpose(vTmp)
            '2020-09 Replace Application.Transpose with Chip Pearson's TransposeArray...
            ' but only if the array contains two dimensions
            If NumberOfArrayDimensions(vTmp) = 2 Then
                TransposeArray vTmp, vTransposeResult
                rng.Resize(UBound(vTmp), 1).Value = vTransposeResult
                Erase vTransposeResult
            Else
                rng.Resize(UBound(vTmp), 1).Value = Application.Transpose(vTmp)
            End If
                
            'if we have reached the total number of samples for that region, then move to the next region
            If iDataColumnCounter = iTotNumSamplingsRequested Then
                'Increment to next region
                iMultiCounter = iMultiCounter + 1
                'Reset the counter
                iDataColumnCounter = 0
            End If
            
        ElseIf iDoMultiple = BuildSheetsByQueryTerm Then
            For i = 1 To iQTCount
                'Add a column head for this data extraction for each data worksheet
                wksGTData(i).Cells(1, iDataColumnCounter + 1).Value = "Sampling " & iDataColumnCounter
                lDateCounter = Application.match(Application.EDate(StartDate, 0) _
                                               , wksGTData(i).Range(wksGTData(i).Cells(1, 1) _
                                               , wksGTData(i).Cells(iMaxRow, 1)), 0)
                Set rng = wksGTData(i).Cells(lDateCounter, iDataColumnCounter + 1)
                
                'Extract the right column from v so that it can be written to the worksheet
                Call TransferOneDimensionFromArray(vSourceArr:=V _
                                                     , vDestinationArr:=vTmp _
                                                     , iDimOneOrDimTwo:=2 _
                                                     , lWhichItem:=i + 1 _
                                                     , lStartPoint:=2 _
                                                     , lEndPoint:=UBound(V, 2) _
                                                     , lDestArrBase:=1)
                
                'Write the data to the worksheet
                '2020-09: Not using TransposeArray for a single dimension array
                rng.Resize(UBound(vTmp), 1).Value = Application.Transpose(vTmp)
            
            Next i
        
'        ElseIf iDoMultiple = BuildSheetsByBoth Then
            'Do nothing, as this is not be allowed by the on-sheet error checking
        End If
        
        'Count the query towards the running count
        iQueryCounter = iQueryCounter + 1
        
        'If the query is being processed faster than the quota, then pause
        iQPS = iQPS + 1
        If iQPS = iMaxQPS Then
            If Now < TimeValue("0:00:01") + TimePerQuery Then Application.Wait (Now + TimeValue("0:00:01"))
            'Reset iQPS
            iQPS = 0
        End If
        
        'As an additional safeguard, if the number of queries is at the total,
        ' then save the workbook. This prevents data loss,
        ' and introduces an additional delay so that the query rate is not exceeded
        If iQueryCounter = iMaxQueriesP100S Then
            'Reset the counter
            iQueryCounter = 0
            
            'Save the workbook
            'wbkNew.Save
            SaveWithErrorHandling iSaveOrSaveAs:=1 _
                    , wbk:=wbkNew

            'Application.Wait (Now + TimeValue("0:00:01"))
        End If
        
    Next j
    
    'When all the data have been extracted, complete the reporting of the results workbook (wbkNew)
    CompleteReporting iMultipleCount:=iMultipleCount _
                    , iMaxRow:=iMaxRow _
                    , wbkNew:=wbkNew _
                    , wksGTData:=wksGTData _
                    , wksSummary:=wksSummary _
                    , wksMainSummary:=wksMainSummary _
                    , wksQuerySpec:=wksQuerySpec _
                    , TimeProcessStart:=TimeProcessStart _
                    , iDoMultiple:=iDoMultiple _
                    , sMultipleTitle:=sMultipleTitle _
                    , sMultipleColHeads:=sMultipleColHeads _
                    , vQueryList:=vQueryList _
                    , lCols:=UBound(vURLArray) + 1 _
                    , iTotNumSamplingsRequested:=iTotNumSamplingsRequested _
                    , iSamplingsAchieved:=j - 1

'    For i = 1 To 20
'    Beep
'    Next i

End Sub

Sub CompleteReporting(ByRef iMultipleCount As Integer _
                            , ByRef iMaxRow As Integer _
                            , ByRef wbkNew As Workbook _
                            , ByRef wksGTData() As Worksheet _
                            , ByRef wksSummary() As Worksheet _
                            , ByRef wksMainSummary As Worksheet _
                            , ByRef wksQuerySpec As Worksheet _
                            , ByRef TimeProcessStart As Single _
                            , ByRef iDoMultiple As BuildSheets _
                            , ByRef sMultipleTitle As String _
                            , ByRef sMultipleColHeads() As String _
                            , ByRef vQueryList() As Variant _
                            , ByRef lCols As Long _
                            , ByRef iTotNumSamplingsRequested As Integer _
                            , ByRef iSamplingsAchieved As Integer)
'This procedure takes variables from DrawSampleDataFromGoogle and uses them to complete the
' Query specification (auditing) sheet (wksQuerySpec), as well as to
' add a chart plotting the values listed on the Main summary sheet (wksMainSummary)
' Which will be the only summary sheet if only a single query term and a single geographical region are specified
    
    Dim i As Integer                'General purpose counter
    Dim lCellsWithZeroTot As Long   'Counts if cells containing zero values were found, totalling across all sheets
    Dim lCellsWithZero1S As Long    'Counts if cells containing zero values were found in one specific sheet
    Dim TimeProcessEnd As Single    'Used to time how long the extraction process takes
    Dim TimeProcessTotal As Single  'Used to time how long the extraction process takes
    Dim sChartTitle As String       'Creates a title for the chart
    Dim RngTmp As Range             'A temporary range object used for the creating of worksheet names
    
    'All samplings have been requested, so clear the status bar reporting
    Application.StatusBar = vbNullString
    
    'Tidy up the reporting
    For i = 1 To iMultipleCount
        With wksGTData(i)
            .Activate
            .Range(.Cells(1, 1), .Cells(1, lCols)).Font.Bold = True
            .Range(.Cells(2, 2), .Cells(iMaxRow, lCols)).Style = "Normal"  'NumberFormat = "General"
            .UsedRange.Columns.AutoFit
            'Check that no columns are too wide after the autofitting
            DialBackColumnWidths wks:=wksGTData(i)
            
            DoFreezePanes wks:=wksGTData(i), sRng:="B2"
            
            'Delete all cells with zero data, counting how many are found in the process
            lCellsWithZero1S = Application.WorksheetFunction.CountIf(.UsedRange, 0)
            If lCellsWithZero1S > 0 Then
                lCellsWithZeroTot = lCellsWithZeroTot + lCellsWithZero1S
                'Do the find and replace twice, once to mark the cells (yellow), and then another time to clear them
                With Application.ReplaceFormat.Interior
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                .Cells.Replace what:="0", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=True
                .Cells.Replace what:="0", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                .Cells(1, 1).AddComment lCellsWithZero1S & " cells with zero values were cleared."
            End If
            On Error GoTo 0
        End With
        'Remove the find & replace modifications
        With Application
            .FindFormat.Clear
            .ReplaceFormat.Clear
            If Not ActiveCell Is Nothing Then
                Cells.Find what:=vbNullString _
                         , After:=ActiveCell _
                         , LookIn:=xlFormulas _
                         , LookAt:=xlPart _
                         , SearchOrder:=xlByRows _
                         , SearchDirection:=xlNext _
                         , MatchCase:=False _
                         , SearchFormat:=False
                Cells.Replace what:=vbNullString _
                            , Replacement:=vbNullString _
                            , ReplaceFormat:=False
            End If
        End With
        
        'Calculate the summary
        'Also add worksheet (not workbook) names to the summary
        With wksSummary(i)
            .Range(.Cells(1, 2), .Cells(1, 16)).Value = _
                Array("N", "Min", "Max", "Range", "Median", "Mean", "StdDev", "Coeff of Variation", "1% Margin of Error", "MoE as % of mean", "99% LCL", "99% UCL", "N needed (1%MoE, 1%CI)" _
                    , IIf(Len(sMultipleTitle) > 0, sMultipleTitle, "Query term") & ":", sMultipleColHeads(i))
            .Range(.Cells(1, 1), .Cells(1, 14)).Font.Bold = True
            .Cells(1, 15).Font.Italic = True
'            .Range(.Cells(2, 2), .Cells(iMaxRow, 2)).FormulaR1C1 = "=COUNT('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")"
'            .Range(.Cells(2, 3), .Cells(iMaxRow, 3)).FormulaR1C1 = "=IF(RC2>0,MIN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 4), .Cells(iMaxRow, 4)).FormulaR1C1 = "=IF(RC2>0,MAX('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 5), .Cells(iMaxRow, 5)).FormulaR1C1 = "=IFERROR(RC[-1]-RC[-2]," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 6), .Cells(iMaxRow, 6)).FormulaR1C1 = "=IFERROR(MEDIAN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 7), .Cells(iMaxRow, 7)).FormulaR1C1 = "=IFERROR(AVERAGE('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 8), .Cells(iMaxRow, 8)).FormulaR1C1 = "=IFERROR(STDEV.S('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 9), .Cells(iMaxRow, 9)).FormulaR1C1 = "=IFERROR(RC8/RC7," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 10), .Cells(iMaxRow, 10)).FormulaR1C1 = "=IFERROR(CONFIDENCE.NORM(0.01,RC8,RC2)," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 11), .Cells(iMaxRow, 11)).FormulaR1C1 = "=IFERROR(RC9/RC7," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 11), .Cells(iMaxRow, 11)).NumberFormat = "#0.0000000%"
'            .Range(.Cells(2, 12), .Cells(iMaxRow, 12)).FormulaR1C1 = "=IFERROR(RC7-RC9," & sQuote & sQuote & ")"
'            .Range(.Cells(2, 13), .Cells(iMaxRow, 13)).FormulaR1C1 = "=IFERROR(RC7+RC9," & sQuote & sQuote & ")"
            
''            ' RefersToR1C1:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlR1C1)
''            Set RngTmp = .Range(.Cells(2, 2), .Cells(iMaxRow, 2))
''            With .Range(.Cells(2, 2), .Cells(iMaxRow, 2))
''                .FormulaR1C1 = "=COUNT('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")"
''                'Debug.Print wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlR1C1)
''                wksSummary(i).Names.Add Name:="_N", RefersTo:=RngTmp    '"'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            Set RngTmp = .Range(.Cells(2, 3), .Cells(iMaxRow, 3))
''            With .Range(.Cells(2, 3), .Cells(iMaxRow, 3))
''                .FormulaR1C1 = "=IF(RC2>0,MIN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_Min", RefersTo:=RngTmp  '"'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 4), .Cells(iMaxRow, 4))
''                .FormulaR1C1 = "=IF(RC2>0,MAX('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_Max", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 5), .Cells(iMaxRow, 5))
''                .FormulaR1C1 = "=IFERROR(RC[-1]-RC[-2]," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_Range", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 6), .Cells(iMaxRow, 6))
''                .FormulaR1C1 = "=IFERROR(MEDIAN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_Median", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 7), .Cells(iMaxRow, 7))
''                .FormulaR1C1 = "=IFERROR(AVERAGE('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_Mean", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 8), .Cells(iMaxRow, 8))
''                .FormulaR1C1 = "=IFERROR(STDEV.S('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_SD", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 9), .Cells(iMaxRow, 9))
''                .FormulaR1C1 = "=IFERROR(RC8/RC7," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_CV", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 10), .Cells(iMaxRow, 10))
''                .FormulaR1C1 = "=IFERROR(CONFIDENCE.NORM(0.01,RC8,RC2)," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_MoE", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 11), .Cells(iMaxRow, 11))
''                .FormulaR1C1 = "=IFERROR(RC9/RC7," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_MoE_percent_Mean", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''                .NumberFormat = "#0.0000000%"
''            End With
''            With .Range(.Cells(2, 12), .Cells(iMaxRow, 12))
''                .FormulaR1C1 = "=IFERROR(RC7-RC9," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_LCL99", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 13), .Cells(iMaxRow, 13))
''                .FormulaR1C1 = "=IFERROR(RC7+RC9," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_UCL99", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
''            With .Range(.Cells(2, 14), .Cells(iMaxRow, 14))
''                .FormulaR1C1 = "=IFERROR(((NORM.INV(1-0.99/2,0,1)*RC8)/0.01)^2," & sQuote & sQuote & ")"
''                wksSummary(i).Names.Add Name:="_N_Needed", RefersTo:="'" & wksSummary(i).Name & "'!" & .Address(ReferenceStyle:=xlA1)
''            End With
            Set RngTmp = .Range(.Cells(2, 2), .Cells(iMaxRow, 2))
            RngTmp.FormulaR1C1 = "=COUNT('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")"
            wksSummary(i).Names.Add Name:="_N", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 3), .Cells(iMaxRow, 3))
            RngTmp.FormulaR1C1 = "=IF(RC2>0,MIN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_Min", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 4), .Cells(iMaxRow, 4))
            RngTmp.FormulaR1C1 = "=IF(RC2>0,MAX('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_Max", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 5), .Cells(iMaxRow, 5))
            RngTmp.FormulaR1C1 = "=IFERROR(RC[-1]-RC[-2]," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_Range", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 6), .Cells(iMaxRow, 6))
            RngTmp.FormulaR1C1 = "=IFERROR(MEDIAN('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_Median", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 7), .Cells(iMaxRow, 7))
            RngTmp.FormulaR1C1 = "=IFERROR(AVERAGE('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_Mean", RefersTo:=RngTmp
            'xx check these calculations
            Set RngTmp = .Range(.Cells(2, 8), .Cells(iMaxRow, 8))
            RngTmp.FormulaR1C1 = "=IFERROR(STDEV.S('" & wksGTData(i).Name & "'!RC2:RC" & lCols & ")," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_SD", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 9), .Cells(iMaxRow, 9))
            RngTmp.FormulaR1C1 = "=IFERROR(RC8/RC7," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_CV", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 10), .Cells(iMaxRow, 10))
            RngTmp.FormulaR1C1 = "=IFERROR(CONFIDENCE.NORM(0.01,RC8,RC2)," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_MoE", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 11), .Cells(iMaxRow, 11))
            RngTmp.FormulaR1C1 = "=IFERROR(RC10/RC7," & sQuote & sQuote & ")"
            RngTmp.NumberFormat = "#0.0000000%"
            wksSummary(i).Names.Add Name:="_MoE_percent_Mean", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 12), .Cells(iMaxRow, 12))
            RngTmp.FormulaR1C1 = "=IFERROR(RC7-RC10," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_LCL99", RefersTo:=RngTmp
            
            Set RngTmp = .Range(.Cells(2, 13), .Cells(iMaxRow, 13))
            RngTmp.FormulaR1C1 = "=IFERROR(RC7+RC10," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_UCL99", RefersTo:=RngTmp

            Set RngTmp = .Range(.Cells(2, 14), .Cells(iMaxRow, 14))
            RngTmp.FormulaR1C1 = "=IFERROR(ROUNDUP(((NORM.INV(1-0.01/2,0,1)*RC8)/(0.01*RC7))^2,0)," & sQuote & sQuote & ")"
            'xx
            'Possibly consider 5% MoE?
            'RngTmp.FormulaR1C1 = "=IFERROR(((NORM.INV(1-0.01/2,0,1)*RC8)/(0.05*RC7))^2," & sQuote & sQuote & ")"
            wksSummary(i).Names.Add Name:="_N_Needed", RefersTo:=RngTmp
            
            .Range(.Cells(1, 1), .Cells(iMaxRow, 14)).AutoFilter
            .Range(.Cells(1, 1), .Cells(iMaxRow, 15)).Columns.AutoFit
        
            'Check that no columns are too wide after the autofitting
            DialBackColumnWidths wks:=wksSummary(i)
            
            DoFreezePanes wks:=wksSummary(i), sRng:="A2"

            If lCellsWithZero1S > 0 Then
                With .Cells(iMaxRow + 1, 1)
                    .Value = lCellsWithZero1S & " cells with zero values were cleared."
                    With .Interior
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End With
            End If
        End With    'wksSummary
    Next i  '=1 to iQTCount
    
    'Build the chart title
    sChartTitle = "Geographic location: "
    Select Case sMultipleTitle
    Case "Region"
        sChartTitle = sChartTitle & fvReturnNameValue("Country", False) & "--All regions"
    Case Else
        If fvReturnNameValue("GeographicLevel", False) = "Worldwide" Then
            sChartTitle = sChartTitle & "Worldwide"
        Else
            sChartTitle = sChartTitle & fvReturnNameValue("Country", False)
        End If
        If fvReturnNameValue("GeographicLevel", False) = "Region" Then _
            sChartTitle = sChartTitle & "--" & fvReturnNameValue("Region", False)
    End Select
    
    sChartTitle = sChartTitle & vbLf _
        & "Time range: " _
        & IIf(fvReturnNameValue("DateResolution") = "Day", "Daily", fvReturnNameValue("DateResolution") & "ly") & " from " _
        & Format(fvReturnNameValue("StartDate"), "dd mmm yyyy") & " to " _
        & Format(fvReturnNameValue("EndDate"), "dd mmm yyyy")
    
    Select Case sMultipleTitle
    'The specification below is explained by the restriction of doing either a multi-region or a multi-term search, but not both
    Case vbNullString
        sChartTitle = sChartTitle & vbLf & "Term: " & sMultipleColHeads(1)
    Case "Region"
        sChartTitle = sChartTitle & vbLf & "Term: " & fvReturnNameValue("SearchTerm01", False)
    Case "Query Term"
        sChartTitle = sChartTitle & vbLf & "Terms: " & "All terms"
    End Select
    'Check that the chart title is less than the limit of 255 characters
    If Len(sChartTitle) > 255 Then _
    sChartTitle = Left(sChartTitle, 252) & Chr(133)
    
    'Now add the chart
    AddChartToOutput wbkNew:=wbkNew _
                   , wksMainSummary:=wksMainSummary _
                   , wksSummary:=wksSummary _
                   , iMultipleCount:=iMultipleCount _
                   , iMaxRow:=iMaxRow _
                   , sChartTitle:=sChartTitle _
                   , iDoMultiple:=iDoMultiple _
                   , sMultipleColHeads:=sMultipleColHeads _
                   , vQueryList:=vQueryList

''    'Create a worksheet which combines all the successive samplings into complete samples
''    'This will not be needed in the final version that goes into the Google Trends extraction tool, as wksSummary already creates the one mean series from all the samplings, and that is sufficient.
''    'But for the article on how many samples I need, I need to create a worksheet that combines all the samplings
''
''    CreateCombinedWorksheet wbkNew, wksCompleteSamples, iMaxRow, i
    
    'Write the last information (time complete, # samples) to the auditing sheet (wksQuerySpec)
    With wksQuerySpec
        .Cells(4, 8).Value = Now
        If bQuotaExceeded Or bOtherHTTPError Or iSamplingsAchieved < iTotNumSamplingsRequested Then
            .Cells(17, 5).Value = "(" & Round(((iSamplingsAchieved - 1) / 2), 0) & " achieved)"
            .Parent.Names.Add Name:="_NSamplesAchieved", RefersTo:="=" & Round(((iSamplingsAchieved - 1) / 2), 0)
            .Parent.Worksheets("Sampling summary").Cells(4, 1).AddComment.Text Text:="Only " & Round(((iSamplingsAchieved - 1) / 2), 0) & " samples were achieved"
            
            .Cells(18, 5).Value = "(" & iSamplingsAchieved - 1 & " achieved)"
        End If
    End With
    Dim sLogOld As String
    sLogOld = sLogEntry
    sLogEntry = sLogEntry & "," & "Samplings achieved:" & iSamplingsAchieved & ","
    sLogEntry = sLogEntry & "Samples achieved:" & Round(((iSamplingsAchieved - 1) / 2), 0) & ","
    sLogEntry = sLogEntry & "Extraction completed:" & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    UpdateLogEntry sLogEntryOld:=sLogOld, sLogEntryNew:=sLogEntry
    sLogOld = vbNullString
    sLogEntry = vbNullString
    
    'Save the workbook for the last time
    'wbkNew.Save
    SaveWithErrorHandling iSaveOrSaveAs:=1 _
                        , wbk:=wbkNew
    
    'Calculate the duration of the process and report to the user
    TimeProcessEnd = Timer
    TimeProcessTotal = IIf(TimeProcessEnd > TimeProcessStart, TimeProcessEnd - TimeProcessStart, TimeProcessEnd + (86400 - TimeProcessStart)) / 86400
    
    'Update the Target file message to show that this file now exists
    With Sheet3
        .Range("QueriesThisSession").Value = fvReturnNameValue("QueriesThisSession", False) + fvReturnNameValue("QueriesNeeded", False)
        .Range("DataTarget").Value = .Range("DataTarget").Value 'Reset this so that it will be recalculated
        .Range("TargetFileMessage1").Calculate
    End With
    
    If bShowCompletionMsgBoxes Then
        MsgBox iSamplingsAchieved & " Samplings (of " & iTotNumSamplingsRequested * IIf(iDoMultiple <= 1, 1, iMultipleCount) _
            & ") for " & fvReturnNameValue(sName:="Samples", bCheckForActiveWorkbook:=False) _
            & " samples have been extracted from Google Trends for your query, and saved as:" _
            & vbCrLf & wbkNew.FullName _
            & vbCrLf & lCellsWithZeroTot & " cells with zero values were cleared." _
            & vbCrLf & "The extraction process took " & Format(TimeProcessTotal, "hh:mm:ss") _
            , vbInformation + vbOKOnly, "Google Trends extraction complete"

        wbkNew.Activate
        
        'Turn on Cut/CopyPaste for the new workbook
        TurnOnPaste
        
        EndGracefully
    Else
        wbkNew.Close SaveChanges:=True
    End If
    
End Sub

Sub SetUpQuerySpecSheet(ByRef wksQuerySpec As Worksheet _
                      , ByRef vQueryList() As Variant _
                      , ByRef vRegions() As Variant _
                      , ByRef vURLs() As Variant)
'Create a worksheet that summarises the query specification and the extraction process
' This allows later auditing of what was done.

    Dim rng As Range
    Dim iLastRow As Integer
    Dim vTransposeResult() As Variant           'Created this 2020-09 to use in the TransposeArray output
    Const iTempLastRow As Integer = 23   'Use this to define a last row (might be exceeded if a large number of query terms is requested)
    
    With wksQuerySpec
        .Name = "Query specification"
        
        sLogEntry = "Google Trends Extended Data for Health API extraction"
        
        With .Cells(1, 1)
            .Value = sLogEntry
            .Style = "Heading 1"
        End With
        sLogEntry = sLogEntry & " (" & fCurrentVersionNumber & "),"
        
        'List the query terms
        With .Cells(2, 1)
            .Value = "Query terms"
            .Style = "Heading 2"
        End With
        
        'Write the query list to the worksheet
        iLastRow = Application.WorksheetFunction.Max(iTempLastRow, UBound(vQueryList) + 3)
        Set rng = .Cells(3, 1)
        '2020-09: Not using TransposeArray for a single dimension array
        rng.Resize(UBound(vQueryList), 1).Value = Application.Transpose(vQueryList)
        sLogEntry = sLogEntry & "Query Terms:" & fReturnQueryListAsOneString(sWorksheetPrefix:=vbNullString) & ","
        
        'Copy the date specification
        CopyRange rngTarget:=.Range(.Cells(2, 3), .Cells(5, 4)) _
                , rngSource:=Sheet3.Range("CompleteDateRangeSpecification")
        .Cells(2, 3).Style = "Heading 2"
        
        sLogEntry = sLogEntry & "Date resolution:" & fvReturnNameValue("DateResolution") & ","
        sLogEntry = sLogEntry & "Start Date:" & Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") & ","
        sLogEntry = sLogEntry & "End Date:" & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") & ","
        
        'Copy the geographic specification
        CopyRange rngTarget:=.Range(.Cells(7, 3), .Cells(11, 4)) _
                , rngSource:=Sheet3.Range("CompleteLocationRangeSpecification")
        Select Case fvReturnNameValue("GeographicLevel")
        Case "Worldwide"
            .Range(.Cells(9, 4), .Cells(10, 4)).ClearContents
        Case "Country"
            .Cells(10, 4).ClearContents
        Case "Region"
            'If all regions are to be queried (i.e., Region was left blank), then write all regions to the sheet
            If UBound(vRegions) > 1 Then
                Set rng = .Cells(10, 4)
                '2020-09: Not replacing this application.transpose because it is a single dimension, string array
                rng.Resize(1, UBound(vRegions)).Value = Application.Transpose(vRegions)
            End If
        End Select
        .Cells(7, 3).Style = "Heading 2"

        sLogEntry = sLogEntry & "Geographic Level:" & fvReturnNameValue("GeographicLevel") & ","
        sLogEntry = sLogEntry & "Country:" & fvReturnNameValue("Country") & ","
        sLogEntry = sLogEntry & "Region:" & fvReturnNameValue("Region") & ","
        sLogEntry = sLogEntry & "Descriptor:" & fvReturnNameValue("Descriptor") & ","
        
        'Copy the sampling specification
        CopyRange rngTarget:=.Range(.Cells(15, 3), .Cells(17, 4)) _
                , rngSource:=Sheet3.Range("CompleteSamplingRangeSpecification")
        .Cells(15, 3).Style = "Heading 2"
        
        .Parent.Names.Add Name:="_NPeriods", RefersTo:=.Cells(16, 4)
        .Parent.Names.Add Name:="_NSamples", RefersTo:=.Cells(17, 4)
        
        sLogEntry = sLogEntry & "Periods:" & fvReturnNameValue("Periods") & ","
        sLogEntry = sLogEntry & "Samples:" & fvReturnNameValue("Samples") & ","
        
        .Cells(18, 3).Value = "Samplings needed:"
        'This formula calculates the number of samplings needed according to the sampling strategy I devised
        .Cells(18, 4).FormulaR1C1 = "=R[-1]C*2-1"
        
        sLogEntry = sLogEntry & "Samplings needed:" & fvReturnNameValue("Samples") * 2 - 1 & ","
        
        'Set the date and time of extraction
        With .Cells(2, 7)
            .Value = "Date and time of extraction:"
            .Style = "Heading 2"
        End With
        .Cells(3, 7).Value = "Start:"
        With .Cells(3, 8)
            .Value = Now()
            .NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End With
        
        .Cells(4, 7).Value = "End:"
        .Cells(4, 8).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Cells(5, 7).Value = "Duration:"
        .Cells(5, 8).FormulaR1C1 = "=R[-1]C-R[-2]C"
        .Cells(5, 8).NumberFormat = "[hh]:mm:ss"
                
        With .Range(.Cells(2, 1), .Cells(iLastRow, 8))
            .HorizontalAlignment = xlLeft
            .Columns.AutoFit
        End With
        'Check that no columns are too wide after the autofitting
        DialBackColumnWidths wks:=wksQuerySpec
        
        'Log the first URL to show what was being searched [This is done after fitting the columns]
        With .Cells(20, 3)
            .Value = "Sample URL:"
            .Style = "Heading 2"
        End With
        'Do not store the API Key in the logged URL
        '.Cells(20, 4).value = Left(vURLs(1), InStr(1, vURLs(1), "&key=") - 1)
        .Cells(20, 4).Value = Left(vURLs(1), InStr(1, vURLs(1), "&key=") + 4) & sReqStrGTEH_APIKey & sReqStrGTEH_JSON_Request
        
        sLogEntry = sLogEntry & "Sample URL:" & Left(vURLs(1), InStr(1, vURLs(1), "&key=") + 4) & sReqStrGTEH_APIKey & sReqStrGTEH_JSON_Request & ","
        
        'Sign the workbook
        .Cells(iLastRow, 1).Value = Replace(sSignaturePhrase _
            , "Google Trends Information Extraction Tool" _
            , "Google Trends Information Extraction Tool" & " (" & fCurrentVersionNumber & ")" _
            , , , vbTextCompare)
        
        'XX Add web link
        
        Sheet7.Range("GPLv3License").Cells(1, 1).Copy .Cells(iLastRow + 1, 1)
        Sheet7.Range("GPLv3License").Cells(1, 2).Copy .Cells(iLastRow + 2, 1)
        '.Cells(iLastRow + 1, 1).value = "XX Add CC license"
        
        .Activate
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
        End With
    
        sLogEntry = sLogEntry & "Target file:" & fvReturnNameValue("DataTarget")
        sLogEntry = sLogEntry & "Extraction started:" & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        WriteLogEntry iLogEvent:=LogEventGTe, sLogEntry:=sLogEntry
    
        'Do a dump of the raw query specification information for robust later retrieval
        'This writes the values to a worksheet in a less formatted version than the above code
        WriteQSvaluesToFile wbkTarget:=wksQuerySpec.Parent _
                          , wksSpecToStore:=Sheet3
    
    End With 'wksQuerySpec
End Sub
'''xx
Sub TestSetUpSamplingSummarySheet()
'c:\Users\JacquesRaubenheimer\Downloads\(facebook)(youtube)(weather)(cricket)(superannuation);AU;2004-01-01--2018-12-31(Yearly);15 samples;2019-01-24.xlsm
    'Application.EnableEvents = False
    StartSmoothly
    Dim w() As Worksheet
    Dim s() As String
    ReDim w(1 To 5)
    ReDim s(1 To 5)
    Set w(1) = ActiveWorkbook.Worksheets("Summary-facebook")
    Set w(2) = ActiveWorkbook.Worksheets("Summary-youtube")
    Set w(3) = ActiveWorkbook.Worksheets("Summary-weather")
    Set w(4) = ActiveWorkbook.Worksheets("Summary-cricket")
    Set w(5) = ActiveWorkbook.Worksheets("Summary-superannuation")
    s(1) = "Summary-facebook"
    s(2) = "Summary-youtube"
    s(3) = "Summary-weather"
    s(4) = "Summary-cricket"
    s(5) = "Summary-superannuation"
    
    SetUpSamplingSummarySheet ActiveWorkbook.Worksheets("Sheet1"), 5, w, s
    
    EndGracefully
    'Application.EnableEvents = True
End Sub


Sub SetUpSamplingSummarySheet(ByRef wksSS As Worksheet _
                      , ByRef iMultipleCount As Integer _
                      , ByRef wksSmy() As Worksheet _
                      , ByRef sColHeads() As String)
'Added version 2.0.0
'Create a worksheet that uses formulas to guage the adequacy of the number of samples specified in estimating the true values
    
    Const QTHeadRow As Integer = 1
    Const iNSamplesHeadRow As Integer = 2
    Const iCVHeadRow As Integer = 9
    Const iNCVRows As Integer = 9
    Const iNPercentilesToShow = 9
    Dim iCVMinRow As Integer
    Dim iCVMaxRow As Integer
    Dim iCVMeanRow As Integer
    Dim iCVMedianRow As Integer
    Dim iModesHeadRow As Integer
    Dim iPercentileSummaryRow As Integer
    Dim iPercentileCalcRow As Integer
    Dim iLastCol As Integer
    Dim iFormulaCol As Integer
    Dim oComment As Comment
    Dim oFC As FormatCondition
    Dim rngNNeed As Range
    
    Const sNSamplesHead As String = "Samples used for estimations"
    Const sNTimePeriods As String = "Time periods"
    Const sNRequested As String = "N requested"
    Const sNMin As String = "Min N achieved"
    Const sNMax As String = "Max N achieved"
    Const sNMedian As String = "Median N achieved"
    Const sNMean As String = "Mean N achieved"

    Const sCVSummaryHead As String = "Summary of Coeff of Var"
    Const sCVMinHead As String = "Min"
    Const sCVMaxHead As String = "Max"
    Const sCVMedianHead As String = "Median"
    Const sCVMeanHead As String = "Mean"
    Const sModesHead As String = "Modes of N-needed"
    Const sPercentileSummaryHead As String = "Summary of N-needed Percentiles"
    Const sPercentileCalcHead As String = "Calculate Percentile for N-Needed"
    
    iLastCol = iMultipleCount * 2 + 1
    
    iCVMinRow = iCVHeadRow + 1
    iCVMaxRow = iCVMinRow + iNCVRows
    iCVMedianRow = iCVMaxRow + 1
    iCVMeanRow = iCVMedianRow + 1
    iModesHeadRow = iCVMeanRow + 1
    iPercentileSummaryRow = iModesHeadRow + 2
    iPercentileCalcRow = iPercentileSummaryRow + 10
    
    With wksSS
        'Set the headers in the first column
        .Cells(iNSamplesHeadRow, 1).Value = sNSamplesHead
        .Cells(iNSamplesHeadRow + 1, 1).Value = sNTimePeriods
        .Cells(iNSamplesHeadRow + 2, 1).Value = sNRequested
        .Cells(iNSamplesHeadRow + 3, 1).Value = sNMin
        .Cells(iNSamplesHeadRow + 4, 1).Value = sNMax
        .Cells(iNSamplesHeadRow + 5, 1).Value = sNMedian
        .Cells(iNSamplesHeadRow + 6, 1).Value = sNMean
        
        .Cells(iCVHeadRow, 1).Value = sCVSummaryHead
        .Cells(iCVMinRow, 1).Value = sCVMinHead
        .Cells(iCVMaxRow, 1).Value = sCVMaxHead
        .Cells(iCVMedianRow, 1).Value = sCVMedianHead
        .Cells(iCVMeanRow, 1).Value = sCVMeanHead
        .Cells(iModesHeadRow, 1).Value = sModesHead
        .Cells(iPercentileSummaryRow, 1).Value = sPercentileSummaryHead
        .Cells(iPercentileCalcRow, 1).Value = sPercentileCalcHead
        
        'Format the first column
        With .Range(.Cells(iNSamplesHeadRow, 1), .Cells(iPercentileCalcRow, 1))
            With .Font
                .Bold = True
                .Italic = True
'                Stop
                .Color = RGB(128, 128, 128) ' xlGray75
'                .TintAndShade = xlGray75
            End With
        End With
        'Un-bold the subheads
        .Range(.Cells(iNSamplesHeadRow + 1, 1), .Cells(iNSamplesHeadRow + 6, 1)).Font.Bold = False
        .Range(.Cells(iCVMinRow, 1), .Cells(iCVMeanRow, 1)).Font.Bold = False
        
        'Format all the header rows
        'Format the N samples header row
        With .Range(.Cells(iNSamplesHeadRow, 1), .Cells(iNSamplesHeadRow, iLastCol))
            With .Font
                .Bold = True
                .Italic = True
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        .Range(.Cells(iNSamplesHeadRow, 2), .Cells(iNSamplesHeadRow, iLastCol)).HorizontalAlignment = xlRight
        
        'Format the Query Terms header row
        With .Range(.Cells(QTHeadRow, 1), .Cells(QTHeadRow, iLastCol))
            .Font.Bold = True
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        
        'Format the Coefficient of Variation header row
        With .Range(.Cells(iCVHeadRow, 1), .Cells(iCVHeadRow, iLastCol))
            With .Font
                .Bold = True
                .Italic = True
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        .Range(.Cells(iCVHeadRow, 2), .Cells(iCVHeadRow, iLastCol)).HorizontalAlignment = xlRight
        
        'Format the Coefficient of Variation maximum row
        With .Range(.Cells(iCVMaxRow, 1), .Cells(iCVMaxRow, iLastCol))
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
        End With
        
        'Format the Coefficient of Variation mean/median rows
        With .Range(.Cells(iCVMedianRow, 1), .Cells(iCVMeanRow, iLastCol))
            .Font.Italic = True
        End With
        
        'Format the Modes of N-needed row
        With .Range(.Cells(iModesHeadRow, 1), .Cells(iModesHeadRow, iLastCol))
            With .Font
                .Bold = True
                .Italic = True
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        .Range(.Cells(iModesHeadRow, 2), .Cells(iModesHeadRow, iLastCol)).HorizontalAlignment = xlRight

        'Format the Percentile Summary Header row
        With .Range(.Cells(iPercentileSummaryRow, 1), .Cells(iPercentileSummaryRow, iLastCol))
            With .Font
                .Bold = True
                .Italic = True
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
        .Range(.Cells(iPercentileSummaryRow, 2), .Cells(iPercentileSummaryRow, iLastCol)).HorizontalAlignment = xlRight
        
        'Format the Percentile Calculation row
        With .Range(.Cells(iPercentileCalcRow, 1), .Cells(iPercentileCalcRow, iLastCol))
            .Font.Bold = True
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End With
    
        'Add the formulas
        For iFormulaCol = 1 To iMultipleCount
            'Add column borders
            With .Range(.Cells(QTHeadRow, iFormulaCol * 2), .Cells(iPercentileCalcRow, iFormulaCol * 2 + 1))
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlHairline
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlHairline
                End With
            End With
            
            '.Cells(QTHeadRow, iFormulaCol * 2).value = Replace(wksSmy(iFormulaCol).Name, "Summary-", vbNullString, , , vbTextCompare)
            .Cells(QTHeadRow, iFormulaCol * 2).Value = sColHeads(iFormulaCol)
            .Range(.Cells(QTHeadRow, iFormulaCol * 2), .Cells(QTHeadRow, iFormulaCol * 2 + 1)).HorizontalAlignment = xlCenterAcrossSelection
            
            'Add the NSamples formulas
            .Cells(iNSamplesHeadRow + 1, iFormulaCol * 2).FormulaR1C1 = "=_NPeriods"
            .Cells(iNSamplesHeadRow + 2, iFormulaCol * 2).FormulaR1C1 = "=_NSamples"
            .Cells(iNSamplesHeadRow + 3, iFormulaCol * 2).FormulaR1C1 = "=MIN('" & wksSmy(iFormulaCol).Name & "'!_N)"
            .Cells(iNSamplesHeadRow + 4, iFormulaCol * 2).FormulaR1C1 = "=MAX('" & wksSmy(iFormulaCol).Name & "'!_N)"
            .Cells(iNSamplesHeadRow + 5, iFormulaCol * 2).FormulaR1C1 = "=MEDIAN('" & wksSmy(iFormulaCol).Name & "'!_N)"
            With .Cells(iNSamplesHeadRow + 6, iFormulaCol * 2)
                .FormulaR1C1 = "=AVERAGE('" & wksSmy(iFormulaCol).Name & "'!_N)"
                .NumberFormat = "# ###.0"
            End With
            'Add the coefficient of variation frequencies
            .Cells(iCVHeadRow, iFormulaCol * 2).Value = "CV"
            .Cells(iCVHeadRow, iFormulaCol * 2 + 1).Value = "N"
            With .Range(.Cells(iCVMinRow, iFormulaCol * 2), .Cells(iCVMeanRow, iFormulaCol * 2))
                .Font.Italic = True
                .NumberFormat = "#.000"
            End With
            'Add a formula to give 10 evenly spaced increments of the CV,
            ' from the Min to the Max as frequency (histogram) bins
            .Range(.Cells(iCVMinRow, iFormulaCol * 2), .Cells(iCVMaxRow, iFormulaCol * 2)).FormulaR1C1 _
                = "=MIN('" & wksSmy(iFormulaCol).Name & "'!_CV)+(ROW()-" & iCVMinRow & _
                ")*((MAX('" & wksSmy(iFormulaCol).Name & "'!_CV)-MIN('" & wksSmy(iFormulaCol).Name & "'!_CV))/" & iNCVRows & ")"
            'Add a formula to count the number of CV values in each bin
            .Range(.Cells(iCVMinRow, iFormulaCol * 2 + 1), .Cells(iCVMaxRow, iFormulaCol * 2 + 1)).FormulaArray = _
                "=FREQUENCY('" & wksSmy(iFormulaCol).Name & "'!_CV,R" & iCVMinRow & "C" & iFormulaCol * 2 & _
                ":R" & iCVMaxRow & "C" & iFormulaCol * 2 & ")"
            'Add the CV Median formula
            .Cells(iCVMedianRow, iFormulaCol * 2).FormulaR1C1 = "=MEDIAN('" & wksSmy(iFormulaCol).Name & "'!_CV" & ")"
            'Add the CV Mean formula
            .Cells(iCVMeanRow, iFormulaCol * 2).FormulaR1C1 = "=AVERAGE('" & wksSmy(iFormulaCol).Name & "'!_CV" & ")"
            .Range(.Cells(iCVMedianRow, iFormulaCol * 2 + 1), .Cells(iCVMeanRow, iFormulaCol * 2 + 1)).FormulaR1C1 = _
                "=COUNTIF('" & wksSmy(iFormulaCol).Name & "'!_CV," & sQuote & "<=" & sQuote & " & RC[-1])"
        
            'Add the N-needed Mode formulas
            'Add a formula to determine the mode of the N needed
            .Cells(iModesHeadRow, iFormulaCol * 2).Value = "Mode"
            .Cells(iModesHeadRow, iFormulaCol * 2 + 1).Value = "N"
            .Cells(iModesHeadRow + 1, iFormulaCol * 2).FormulaR1C1 = "=MODE('" & wksSmy(iFormulaCol).Name & "'!_N_Needed" & ")"
            .Cells(iModesHeadRow + 1, iFormulaCol * 2 + 1).Value = "=COUNTIF('" & wksSmy(iFormulaCol).Name & "'!_N_Needed, R" & iModesHeadRow + 1 & "C" & iFormulaCol * 2 & ")"
        
            'Add the Percentile Summary formulas
            .Cells(iPercentileSummaryRow, iFormulaCol * 2).Value = "Percentile"
            .Cells(iPercentileSummaryRow, iFormulaCol * 2 + 1).Value = "N needed"
            With .Range(.Cells(iPercentileSummaryRow + 1, iFormulaCol * 2), .Cells(iPercentileSummaryRow + iNPercentilesToShow, iFormulaCol * 2))
                '2020-09: Not replacing this application.transpose because it is a single dimension, string array
                .Value = Application.Transpose(Array("0%", "1%", "5%", "25%", "50%", "75%", "95%", "99%", "100%"))
                '.NumberFormat = "0.0%"
            End With
            .Range(.Cells(iPercentileSummaryRow + 1, iFormulaCol * 2 + 1), .Cells(iPercentileSummaryRow + iNPercentilesToShow, iFormulaCol * 2 + 1)).FormulaR1C1 = _
                "=PERCENTILE.INC('" & wksSmy(iFormulaCol).Name & "'!_N_Needed, RC[-1])"
        
            'Add the Percentile Estimation formulas
            With .Cells(iPercentileCalcRow, iFormulaCol * 2)
                .FormulaR1C1 = "=PERCENTRANK.INC('" & wksSmy(iFormulaCol).Name & "'!_N_Needed, R" & iPercentileCalcRow & "C[1])"
                .NumberFormat = "#.00%"
            End With
            
            Set rngNNeed = .Range(.Cells(iPercentileSummaryRow + 1, iFormulaCol * 2 + 1), .Cells(iPercentileSummaryRow + iNPercentilesToShow, iFormulaCol * 2 + 1))
            'If this is the first column, then add a reference value, otherwise add a formula so that all query terms are calculated together
            With .Cells(iPercentileCalcRow, iFormulaCol * 2 + 1)
                
                .Style = "Input"
                
                If iFormulaCol = 1 Then
                    .FormulaR1C1 = "=MEDIAN(R" & iPercentileSummaryRow + 1 & "C" & iFormulaCol * 2 + 1 & ":R" & iPercentileSummaryRow + iNPercentilesToShow & "C" & iFormulaCol * 2 + 1 & ")"
'                    .value = Application.WorksheetFunction.Median(rngNNeed)
                    Set oComment = .AddComment
                    With oComment
                        .Text Text:="Google Trends Data extraction tool:" _
                            & Chr(10) & "Enter the desired N to see the percentile in the N-needed distribution" _
                            & Chr(10) & "(Delete the formula which initially sets it to the median, or the synchronising formulas in columns to the right)."
                        With .Shape
                            .Height = Application.CentimetersToPoints(2.5)
                            .Width = Application.CentimetersToPoints(6)
                        End With
                    End With
                Else
                    .FormulaR1C1 = "=MIN(MAX(R" & iPercentileSummaryRow + 1 & "C" & iFormulaCol * 2 + 1 _
                        & ":R" & iPercentileCalcRow - 1 & "C" & iFormulaCol * 2 + 1 & ")," _
                        & "R" & iPercentileCalcRow & "C3)"
                    .FormatConditions.Add Type:=xlExpression, Formula1:="=RC<>R" & iPercentileCalcRow & "C3"
                    .FormatConditions(1).Interior.Color = vbYellow  '65535
                End If
            End With
        
        Next iFormulaCol

        .Columns.AutoFit
        'Check that no columns are too wide after the autofitting
        DialBackColumnWidths wks:=wksSS
        
        ActiveWindow.DisplayGridlines = False

    End With

End Sub

Function fDoAllErrorChecking() As Boolean
'This function does all the error checking across each of the input areas.
' When any error is found, it alerts the user and ends execution.
' If the end of the function is reached, then theoretically, no errors were encountered.

    'Recalculate the sheet to ensure the on-sheet error checking is up to date
    Sheet3.Calculate

    'Check for any error messages on the worksheet. This is the first-level check, which needs to be passed.
    'There is one error message starting with "Note:"
    'This will happen when the users wants the tool to extract a sample for a single period, and should be allowed
    'So if this is the only error, then continue, otherwise show the error message
    'Also, I have added an 'All good to go' message that reads "All inputs are completed. Data extraction can be commenced."
    If fvReturnNameValue("NErrorMessages") > 1 _
    Or Not (fvReturnNameValue("NErrorMessages") = 1 _
        And (Left(fvReturnNameValue("ErrorDisplay1"), 4) = "Note" _
        Or Left(fvReturnNameValue("ErrorDisplay1"), 3) = "All")) Then
        MsgBox "There are still outstanding input error messages that need to be resolved before the query can be sent." _
            & vbCrLf & "Please attend to these issues and then try again." _
            , vbCritical + vbOKOnly, "Input Error Messages outstanding"
        EndGracefully
    End If

    'Now, even though it is a repetition of the UI error checking on the worksheet,
    ' as a failsafe, each input field is checked in much the same way as on the UI
    
    'These defined names check for completion of various parts of the UI input
    '=CompletedAll
    '=CompletedAPIKey
    '=CompletedCategory
    '=CompletedDataTarget
    '=CompletedDateInputs
    '=CompletedLocation
    '=CompletedSamples
    '=CompletedSearchTerms
    
    'Check that at least one query term exists
    If Not fvReturnNameValue("CompletedSearchTerms") Then
        Range("SearchTerm01").Select
        MsgBox "You must specify at least one Query term!" _
            , vbCritical + vbOKOnly _
            , "No query terms"
        EndGracefully
    End If
    
    'Check that the search location has been specified properly
    If Not fvReturnNameValue("CompletedLocation") Then
        Range("GeographicLevel").Select
        MsgBox "The geographic location is not specified correctly!" _
            , vbCritical + vbOKOnly _
            , "Incorrection geographic specification"
        EndGracefully
    End If
    'Now check all the possible combinations of the geographic location specification
    'First check the level
    'Check that the level specification exists
    If Not Len(fvReturnNameValue("GeographicLevel")) > 0 Then
        Range("GeographicLevel").Select
        MsgBox "You must specify the Geographic Level!" _
            , vbCritical + vbOKOnly _
            , "No Geographic Level specification"
        EndGracefully
    End If
    'Check that the level is in the allowable list
    If Not fIsInListNamedRange(fvReturnNameValue("GeographicLevel"), "GeoLevels", False) Then
        Range("GeographicLevel").Select
        MsgBox "The value '" & fvReturnNameValue("GeographicLevel") & "' specified for the Geographic level is not a valid entry!" _
            , vbCritical + vbOKOnly _
            , "Incorrect Geographic level specification"
        EndGracefully
    End If
    
    Select Case fvReturnNameValue("GeographicLevel")
'    Case "Worldwide"
'        'Do nothing further
'    Case Else
    Case "Country", "Region"
        'If the level is Country, check the country value as well
        'If the level is Region, check the country and region values as well
        'So first we check the country, because that must be checked regardless
        'Check that the Country is specified
        If Not Len(fvReturnNameValue("Country")) > 0 Then
            Range("Country").Select
            MsgBox "You must specify the Country!" _
                , vbCritical + vbOKOnly _
                , "No Country specification"
            EndGracefully
        End If
        'Check that the Country is in the allowable list
        If Not fIsInListNamedRange(fvReturnNameValue("Country"), "CountryNames", False) Then
            Range("Country").Select
            MsgBox "The value '" & fvReturnNameValue("Country") & "' specified for the Country is not a valid entry!" _
                , vbCritical + vbOKOnly _
                , "Incorrect Country specification"
            EndGracefully
        End If
        
        'Then if the level is region, we check that too
        If fvReturnNameValue("GeographicLevel") = "Region" Then
            'Originally, the user had to specify a region, but this check is now disabled,
            ' as the user is allowed to leave the region blank, and so sample all regions for the specified country
            ' If the country has many regions, the UI will warn them about the impracticality of this,
            ' given the 2000 data-point limit
''            'Check that the Region exists
''            If Not Len(fvReturnNameValue("Region")) > 0 Then
''                Range("Region").Select
''                MsgBox "You must specify the Region!" _
''                    , vbCritical + vbOKOnly _
''                    , "No Region specification"
''                EndGracefully
''            End If
            
            'Check that the Region is in the allowable list
''            If Not fIsInListNamedRange(fvReturnNameValue("Region"), "CountrySubdivisionDynamicList", False) Then
            If Not fIsInListNamedRange(fvReturnNameValue("Region"), "CountrySubdivisionDynamicList", False) And Not fvReturnNameValue("Region") = vbNullString Then
                Range("Region").Select
                MsgBox "The value '" & fvReturnNameValue("Region") & "' specified for the Region is not a valid entry!" _
                    , vbCritical + vbOKOnly _
                    , "Incorrect Region specification"
                EndGracefully
            End If
        End If
    End Select
    
    'Check that the start and end dates are valid dates, and that the start date is after the end date
    'These are the UI error checking formulas:
    '=IF(LEN(EndDate)=0,"",IF(EndDate<StartDate,"Warning! The end date is before the start date. Your time period selection is invalid!",IF(Periods=1,"Note: The end date is within a single " & DateResolution & " of the start date, so no sampling is possible.","")))
    '=IF(LEN(StartDate)=0,"",IF(OR(AND(DateResolution="Year",OR(DAY(StartDate)<>1,MONTH(StartDate)<>1)),AND(DateResolution="Month",DAY(StartDate)<>1),AND(DateResolution="Week",WEEKDAY(StartDate,1)<>1)),"Warning! The Start date is not set to the first day of the " & DateResolution & ". This will generate an incomplete first time period. Consider changing the date.",""))
    '=IF(LEN(EndDate)=0,"",IF(OR(AND(DateResolution="Year",OR(DAY(EndDate)<>31,MONTH(EndDate)<>12)),AND(DateResolution="Month",EndDate<>EOMONTH(EndDate,0)),AND(DateResolution="Week",WEEKDAY(EndDate,1)<>7)),"Warning! The End date is not set to the last day of the " & DateResolution & ". This will generate an incomplete last time period. Consider changing the date.",""))
    
''This is commented out since each one is checked individually below
''    'Check the overall date inputs
''    If Not fvReturnNameValue("CompletedDateInputs") Then
''        Range("DateResolution").Select
''        MsgBox "The date specification is incomplete!" _
''            , vbCritical + vbOKOnly _
''            , "Incomplete date range specification"
''        End
''    End If
    
    'Check that the Date Resolution exists
    If Not Len(fvReturnNameValue("DateResolution")) > 0 Then
        Range("DateResolution").Select
        MsgBox "You must specify the Date frequency!" _
            , vbCritical + vbOKOnly _
            , "No Date Resolution specification"
        EndGracefully
    End If
    'Check that the Date Resolution is in the allowable list
    If Not fIsInListNamedRange(fvReturnNameValue("DateResolution"), "DateFrequencyOptions", False) Then
        Range("DateResolution").Select
        MsgBox "The value '" & fvReturnNameValue("DateResolution") & "' specified for the Date frequency is not a valid entry!" _
            , vbCritical + vbOKOnly _
            , "Incorrect Date frequency specification"
        EndGracefully
    End If
    
    'Check that the Start Date exists
    If Not Len(fvReturnNameValue("StartDate")) > 0 Then
        Range("StartDate").Select
        MsgBox "You must provide a Start date!" _
            , vbCritical + vbOKOnly _
            , "No Start Date provided"
        EndGracefully
    End If
    
    'Check that the Start Date is a valid date
    If Not IsDate(fvReturnNameValue("StartDate")) Then
        Range("StartDate").Select
        MsgBox "The Start Date value of '" & fvReturnNameValue("StartDate") & "' is not a valid date!" _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the Start Date is at the start of the specified time period
    Select Case fvReturnNameValue("DateResolution")
    Case "Week"
        If Weekday(fvReturnNameValue("StartDate")) > 1 Then
            Range("StartDate").Select
            'If the week start is before 3 Jan 2004, then the start week is moved forward to 4 Jan 2004, otherwise it is moved back to the first preceding Sunday
            MsgBox "the Start Date " & Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") _
                & " is not at the start of the week, and this will generate incomplete data for the first week." _
                & vbCrLf & "Please set the Start Date to " _
                & Format(IIf(fvReturnNameValue("StartDate") <= DateSerial(2004, 1, 3), DateSerial(2004, 1, 4) _
                , fvReturnNameValue("StartDate") - Weekday(fvReturnNameValue("StartDate")) + 1), "yyyy-mm-dd") & "." _
                , vbCritical + vbOKOnly, "Incomplete Starting period"
            EndGracefully
        End If
    Case "Month"
        If Day(fvReturnNameValue("StartDate")) > 1 Then
            Range("StartDate").Select
            MsgBox "the Start Date " & Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") _
                & " is not at the start of the month, and this will generate incomplete data for the first month." _
                & vbCrLf & "Please set the Start Date to " & Format(DateSerial(Year(fvReturnNameValue("StartDate")), Month(fvReturnNameValue("StartDate")), 1), "yyyy-mm-dd") & "." _
                , vbCritical + vbOKOnly, "Incomplete Starting period"
            EndGracefully
        End If
    Case "Year"
        If Month(fvReturnNameValue("StartDate")) > 1 Or Day(fvReturnNameValue("StartDate")) > 1 Then
            Range("StartDate").Select
            MsgBox "the Start Date " & Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") _
                & " is not at the start of the year, and this will generate incomplete data for the first year." _
                & vbCrLf & "Please set the Start Date to " & _
                Format(DateSerial(Year(fvReturnNameValue("StartDate")), 1, 1), "yyyy-mm-dd") & "." _
                , vbCritical + vbOKOnly, "Incomplete Starting period"
            EndGracefully
        End If
    End Select
    
    'Check that the Start Date is not before 2004/1/1
    If fvReturnNameValue("StartDate") < DateSerial(2004, 1, 1) Then
        Range("StartDate").Select
        MsgBox "The Start Date is before the 1st of January 2004!" _
            & " There are no Google Trends data before that date." _
            & vbCrLf & "Please set the Start Date to " _
            & Format(IIf(fvReturnNameValue("DateResolution") = "Week", DateSerial(2004, 1, 4), DateSerial(2004, 1, 1)), "yyyy-mm-dd") & "." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the End Date is specified
    If Not Len(fvReturnNameValue("EndDate")) > 0 Then
        Range("EndDate").Select
        MsgBox "You must provide a End date!" _
            , vbCritical + vbOKOnly _
            , "No End Date provided"
        EndGracefully
    End If
    
    'Check that the End Date is a valid date
    If Not IsDate(fvReturnNameValue("EndDate")) Then
        Range("EndDate").Select
        MsgBox "The End Date value '" & fvReturnNameValue("EndDate") & " is not a valid date!" _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the End Date is at the End of the specified time period
    'These variables are needed since determining the end date is harder,
    ' as an end date before or after the current date can be chosen,
    ' and the end date may not extend into two days before the present
    Dim maxDateBef As Date, maxDateAft As Date
    maxDateBef = fSetMaxDateBefore(fvReturnNameValue("EndDate"), fvReturnNameValue("DateResolution"))
    maxDateAft = fSetMaxDateAfter(fvReturnNameValue("EndDate"), fvReturnNameValue("DateResolution"))
    
    'Check that the End Date is not later than two days before the present date
    If fvReturnNameValue("EndDate") > Now() - 2 Then
        Range("EndDate").Select
        MsgBox "The End Date " & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") & " is later than " & Format(Now() - 2, "yyyy-mm-dd") & "!" _
            & " There are no Google Trends Extended data available for dates later than two days prior to the present date." _
            & vbCrLf & "Please set the End Date to " & Format(maxDateBef, "yyyy-mm-dd") & "." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    Select Case fvReturnNameValue("DateResolution")
    Case "Week"
        If Weekday(fvReturnNameValue("EndDate")) < 7 Then
            Range("EndDate").Select
            MsgBox "the End Date " & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") _
                & " is not at the end of the week, and this will generate incomplete data for the last week." _
                & vbCrLf & "Please set the End Date to " _
                & IIf(maxDateBef = maxDateAft, Format(maxDateBef, "yyyy-mm-dd"), Format(maxDateBef, "yyyy-mm-dd") & " or " & Format(maxDateAft, "yyyy-mm-dd")) & "." _
                , vbCritical + vbOKOnly, "Incomplete Ending period"
        End If
    Case "Month"
        If fvReturnNameValue("EndDate") < Application.WorksheetFunction.EoMonth(fvReturnNameValue("EndDate"), 0) Then
            Range("EndDate").Select
            MsgBox "the End Date " & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") _
                & " is not at the end of the month, and this will generate incomplete data for the last month." _
                & vbCrLf & "Please set the End Date to " _
                & IIf(maxDateBef = maxDateAft, Format(maxDateBef, "yyyy-mm-dd"), Format(maxDateBef, "yyyy-mm-dd") & " or " & Format(maxDateAft, "yyyy-mm-dd")) & "." _
                , vbCritical + vbOKOnly, "Incomplete Ending period"
        End If
    Case "Year"
        If fvReturnNameValue("EndDate") < DateSerial(Year(fvReturnNameValue("EndDate")), 12, 31) Then
            Range("EndDate").Select
            MsgBox "the End Date " & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") _
                & " is not at the end of the year, and this will generate incomplete data for the last year." _
                & vbCrLf & "Please set the End Date to " _
                & IIf(maxDateBef = maxDateAft, Format(maxDateBef, "yyyy-mm-dd"), Format(maxDateBef, "yyyy-mm-dd") & " or " & Format(maxDateAft, "yyyy-mm-dd")) & "." _
                , vbCritical + vbOKOnly, "Incomplete Ending period"
        End If
    End Select
    
    'Check that the Start Date is not after the End date
    If fvReturnNameValue("StartDate") > fvReturnNameValue("EndDate") Then
        Range("StartDate").Select
        MsgBox "The Start Date (" & Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") & ") is after the End Date (" & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") & ")!" _
            & vbCrLf & "This is an invalid time specification." _
            & vbCrLf & "Please set the Start Date to before the End Date, or the End Date to after the Start Date." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
    End If

    'Check that the number of samples is specified
    If Not fvReturnNameValue("CompletedSamples") Then
        Range("Samples").Select
        MsgBox "The number of samples has not been specified!" _
            , vbCritical + vbOKOnly _
            , "No sample number"
        EndGracefully
    End If
    
    'Check that the number of samples <= the number of periods
    'This is the UI error checking formula:
    '=IF(Samples>Periods,"Warning! Complete sampling is not possible when Samples > Periods!","")
    'Possible to-do item:
    'This might not be needed if I decide to implement the simpler sampling method when the date specification does not
    ' approach the bounds of the available date range
    If fvReturnNameValue("Samples") > fvReturnNameValue("Periods") Then
        Range("Samples").Select
        MsgBox "You have requested " & fvReturnNameValue("Samples") _
            & " samples, but the time specification allows for only " & fvReturnNameValue("Periods") & " periods!" _
            & vbCrLf & "Complete sampling is not possible when Samples > Periods." _
            & vbCrLf & "Please reduce the number of samples or consider lengthening the time period, or setting a finer date frequency to allow more time periods." _
            & " If the number of samples is not enough, you could consider repeating the sample extraction in 24 hours to gain more samples." _
            , vbCritical + vbOKOnly _
            , "Samples > Periods"
        EndGracefully
    End If

    'Check that the number of periods <=2000
    'This is the UI error checking formula:
    '=IF(Periods>2000,"Warning! The Google API allows a maximum of 2000 data points! Please set the Start date to "&TEXT(EndDate-1999,"yyyy/mm/dd")&" or the End date to "&TEXT(StartDate+1999,"yyyy/mm/dd")&".","")
    If fvReturnNameValue("Periods") > 2000 Then
        'Possible to-do item:
        'Consider a procedure to break longer periods into subgroups, sampling from each, and then combining them
        '(Note that this will not work on the normal API, because of the normalisation, but it will work on the extended API)
        'I should still set the number of samplings to be less than 2000, as 2000 samplings really is a bit excessive
        Range("StartDate").Select
        MsgBox "The number of periods (" & fvReturnNameValue("Periods") & ") specified by the date selection you have made" _
            & "(" & fvReturnNameValue("DateResultion") & " for " & fvReturnNameValue("StartDate") & " to " & fvReturnNameValue("EndDate") & ")" _
            & " is greater than 2000. However, the Google API allows only a maximum of 2000 data points in one request." _
            & vbCrLf & "Please set the End Date to " & Format(fvReturnNameValue("StartDate") + 2000 - 1, "yyyy-mm-dd") _
            & " or the Start Date to " & Format(fvReturnNameValue("EndDate") - 2000, "yyyy-mm-dd") & "." _
            , vbCritical + vbOKOnly _
            , "Periods > 2000"
        EndGracefully
    End If
    
    'Check that the API key can be successfully retrieved
    'This is the UI error checking formula:
    '=IF(LEN(APIKey)=0,"You must specify a file containing the API key.",IF(FileDirCheck(1,APIKey),"","Warning! The file specified for the API key cannot be found!"))
    'First check that a file is specified
    If Not fvReturnNameValue("CompletedAPIKey") Then
        Range("APIKey").Select
        MsgBox "The file containing the Google Trends Extended API key has not been specified!" _
            & vbCrLf & "Please create a text file containing only the key, save it using either a *.txt or *.key file extension, and then double click on the cell to access the file." _
            , vbCritical + vbOKOnly _
            , "No API Key"
        EndGracefully
    End If
    
    'Next check that the file actually still exists
    If Not Len(Dir(fvReturnNameValue("APIKey"), vbNormal)) > 0 Then
        Range("APIKey").Select
        MsgBox "The file specified for the API Key cannot be found!" _
            & vbCrLf & "It may have been deleted or moved." _
            & vbCrLf & "Please double click on the cell to specify the new file location." _
            , vbCritical + vbOKOnly _
            , "API Key file not found"
        EndGracefully
    End If
    
    'There is currently no check that the file actually contains an accurate API key, as this will be tested when the call is sent
    
    'Check that the target file directory exists, and that the file itself does not already exist.
    'This is the UI error checking formula:
    '=IF(LEN(DataTarget)=0,"You must specify a Target file where the data extraction results will be stored.",IF(NOT(FileDirCheck(2,DataTarget)),"Warning! The directory specified for the data extraction file cannot be found!",IF(FileDirCheck(1,DataTarget),"Warning! The Target file specified for the data extraction already exists!","")))
    'First check that a file is specified
    If Not fvReturnNameValue("CompletedDataTarget") Then
        Range("DataTarget").Select
        MsgBox "No file and path name has been specified for the output file!" _
            & vbCrLf & "Please double click on the cell to specify the file name and location for the output file." _
            , vbCritical + vbOKOnly _
            , "No output file"
        EndGracefully
    End If
    
    'Check that the specified directory is accessible
    If Not FileDirCheck(2, fvReturnNameValue("DataTarget")) Then
        Range("DataTarget").Select
        MsgBox "The directory specified for the data extraction file cannot be found!" _
            & vbCrLf & "Please double click on the cell to specify the file name and location for the output file." _
            , vbCritical + vbOKOnly _
            , "Directory not found"
        EndGracefully
    End If
    
    'Next check that the file does not already exist
    ' At the moment, overwriting is not allowed.
    If Len(Dir(fvReturnNameValue("DataTarget"), vbNormal)) > 0 Then
        Range("DataTarget").Select
        MsgBox "The file specified for the output file already exists!" _
            & vbCrLf & "Please double click on the cell to specify a new file for the output." _
            , vbCritical + vbOKOnly _
            , "Output file already exists"
        EndGracefully
    End If
    
    'If this point is reached, and no error check has bombed out, then we assume(!) there are no errors
    fDoAllErrorChecking = True

End Function

Function fSetMaxDateBefore(ByVal d As Date, ByVal sDateRes As String) As Date
'Return a maximum last date that the End Date can be set to, which is before the current End Date
'This function is required because of the complexity of having to deal with the date being <= two days prior to the present, and
'also that the date may not extend into the next time period (e.g., the next year or month)
    
    If d > Now() - 2 Then d = Now() - 2
    
    Select Case sDateRes
    Case "Week"
        fSetMaxDateBefore = d - Weekday(d)
    Case "Month"
        fSetMaxDateBefore = DateSerial(Year(d), Month(d), 1) - 1
    Case "Year"
        fSetMaxDateBefore = DateSerial(Year(d), 1, 1) - 1
    End Select

End Function

Function fSetMaxDateAfter(ByVal d As Date, ByVal sDateRes As String) As Date
'Return a maximum last date that the End Date can be set to, which is after the current End Date
'This function is required because of the complexity of having to deal with the date being <= two days prior to the present, and
'also that the date may not extend into the next time period (e.g., the next year or month)
    
    If d > Now() - 2 Then d = Now() - 2
    
    Select Case sDateRes
    Case "Week"
        fSetMaxDateAfter = d + (7 - Weekday(d))
        If fSetMaxDateAfter > Now() - 2 Then fSetMaxDateAfter = d - Weekday(d)
    Case "Month"
        fSetMaxDateAfter = Application.WorksheetFunction.EoMonth(d, 0)
        If fSetMaxDateAfter > Now() - 2 Then fSetMaxDateAfter = DateSerial(Year(d), Month(d), 1) - 1
    Case "Year"
        fSetMaxDateAfter = DateSerial(Year(d) + 1, 1, 1) - 1
        If fSetMaxDateAfter > Now() - 2 Then fSetMaxDateAfter = DateSerial(Year(d) - 1, 12, 31)
    End Select

End Function

'Private Sub testFIsInList()
'    fIsInListNamedRange "Google Search", "Categories"
'End Sub
Function fIsInListNamedRange(ByVal sValueToCheck As String _
                           , ByVal sRangeName As String _
                           , Optional ByVal bStopOnError As Boolean = True) As Boolean
'Check that sValueToCheck is contained somewhere in the range specified by sRangeName
' Used to confirm that a value in the specification is actually in the allowable list
    
    'Validate the Range Name
    On Error Resume Next
    Dim tmp As String
    tmp = Range(sRangeName).Address
    If Err.Number <> 0 Then
            If bStopOnError Then
            MsgBox "The range name '" & sRangeName & "' does not exist. Please check the coding request." _
                , vbCritical + vbOKOnly, "Invalid range name"
            Err.Clear
            EndGracefully
        Else
            Err.Clear
            fIsInListNamedRange = False
            Exit Function
        End If
    End If
    
    fIsInListNamedRange = Application.WorksheetFunction.match(sValueToCheck, Range(sRangeName), 0) > 0
    If Err.Number <> 0 Then
    'Match produces an error when the item is not found (the validity of the range name has already been checked)
        Err.Clear
        fIsInListNamedRange = False
    End If
    On Error GoTo 0

End Function

Private Sub TransferOneDimensionFromArray(ByRef vSourceArr As Variant _
                                     , ByRef vDestinationArr As Variant _
                                     , ByVal iDimOneOrDimTwo As Integer _
                                     , ByVal lWhichItem As Long _
                                     , Optional ByVal lStartPoint As Long = 0 _
                                     , Optional ByVal lEndPoint As Long = 0 _
                                     , Optional ByVal lDestArrBase As Long = -1)

'Takes an array and extracts all or part of a row (Dimension 1) or column (Dimension 2) of a two-dimensional array
' into a single dimension array
'lStartPoint and lEndPoint indicate which values must be transferred
' (leaving them out takes the lower and upper bound in their stead, i.e., leaving both out will take the whole element).
'iDimOneOrDimTwo must be a value of 1 or 2, indicating whether the row (1) or column (2) (i.e., first or second dimension)
' of the array must be transferred.
'lDestArrBase must be -1, 0, or 1. It does not indicate the actual lower bound of the destination array, but rather:
' -if -1, then lStartPoint is taken as the base of the destination array, and it moves up to lEndPoint.
' -if 0 or 1, then the new array is made base 0 or base 1, and lStartPoint and lEndPoint are "offset" to transfer the values to the right position.
    
    'First check the inputs
    If iDimOneOrDimTwo <> 1 And iDimOneOrDimTwo <> 2 Then
        MsgBox "The dimension for transferring the array is not specified correctly." _
            , vbCritical + vbOKOnly, "Incorrect dimension in TransferOneDimensionFromArray"
        EndGracefully
    End If
    If lDestArrBase < -1 Or lDestArrBase > 1 Then
        MsgBox "The destination array base is not specified correctly." _
            , vbCritical + vbOKOnly, "Incorrect destination array base in TransferOneDimensionFromArray"
        EndGracefully
    End If
    
    Dim r As Long           'Row (Dimension 1) counter
    Dim c As Long           'Column (Dimension 2) counter
    Dim lOffset As Long     'Offset value
    
    'If no starting point is specified, then start at the lower bound
    If lStartPoint = 0 Then lStartPoint = LBound(vSourceArr, iDimOneOrDimTwo)
    'If no end point is specified, then end at the upper bound
    If lEndPoint = 0 Then lEndPoint = UBound(vSourceArr, iDimOneOrDimTwo)
    
    'Calculate the offset
    If lDestArrBase = -1 Then
        lOffset = 0
    Else
        lOffset = lStartPoint - lDestArrBase
    End If
    
    'Dimension the destination array according to the calculated Start, End, and Offset values
    ReDim vDestinationArr(lStartPoint - lOffset To lEndPoint - lOffset)
    
    'Transfer the elements
    If iDimOneOrDimTwo = 1 Then
        For r = lStartPoint To lEndPoint
            vDestinationArr(r - lOffset) = vSourceArr(r, lWhichItem)
        Next r
    ElseIf iDimOneOrDimTwo = 2 Then
        For c = lStartPoint To lEndPoint
            vDestinationArr(c - lOffset) = vSourceArr(lWhichItem, c)
        Next c
    End If
End Sub

Private Sub TestSpeedOfTimeFrameMethods()
'This sub is not used.
' I wrote it to test the quickest method of returning the correct time frame
    Dim sTim As String
    Dim s As Single 'Date
    Dim e As Single 'Date
    Dim a As Integer
    Dim tf As TimeFrame
    Dim lSecspDay As Long
'    lSecspDay = 24 * 60 * 60
    lSecspDay = 86400

    'Determine the time frame (simpler than above)
    s = Timer
        For a = 1 To 10000
            sTim = Application.WorksheetFunction.Choose((a Mod 4) + 1, "Year", "Day", "Week", "Month")
            tf = InStr(1, Left$(fvReturnNameValue("DateResolution"), 1), "DWMY", vbTextCompare)
        Next a
    e = Timer
    Debug.Print "Str:" & e - s & "/" & Format((e - s) / lSecspDay, "hh:mm:ss.s")
    
    s = Timer
        For a = 1 To 10000
            sTim = Application.WorksheetFunction.Choose((a Mod 4) + 1, "Year", "Day", "Week", "Month")
            Select Case sTim
            Case "Day"
                tf = TimeFrameDay
            Case "Week"
                tf = TimeFrameWeek
            Case "Month"
                tf = TimeFrameMonth
            Case "Year"
                tf = TimeFrameYear
            End Select
        Next a
    e = Timer
    Debug.Print "SC:" & e - s & "/" & Format((e - s) / lSecspDay, "hh:mm:ss.s")
End Sub

Sub CreateCombinedWorksheet(ByRef wbk As Workbook _
                          , ByRef wks As Worksheet _
                          , ByVal r As Long _
                          , ByVal c As Long)
'This procedure is not used, but was created during production for testing of the process
'It create a worksheet which combines (collapses) all the successive samplings into complete samples
'In the final version wksSummary already creates the one mean series from all the samplings, and that is sufficient,
' as it simply passes over the empty columns.
    
    Dim wksSamples As Worksheet
    Dim sWksName As String
    sWksName = "'" & wks.Name & "'"

    Set wksSamples = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
    With wksSamples
        .Name = "Samples"
        .Cells(1, 1).Value = "Date"
        .Range(.Cells(2, 1), .Cells(r, 1)).NumberFormat = "YYYY-MM"
        .Range(.Cells(2, 1), .Cells(r, 2)).FormulaR1C1 = "=" & sWksName & "!RC"
        .Range(.Cells(1, 2), .Cells(1, c / 2)).FormulaR1C1 = "=" & sQuote & "Sample " & sQuote & "&COLUMN()-1"
        .Range(.Cells(2, 3), .Cells(r, c / 2)).FormulaR1C1 = "=MAX(INDEX(" & sWksName & "!R1C1:R" & r & "C" & c & ",ROW(),COLUMN()*2-3),INDEX(" & sWksName & "!R1C1:R" & r & "C" & c & ",ROW(),COLUMN()*2-2))"
    End With
End Sub
