Attribute VB_Name = "mGTWebFunctions"
'Google Trends Extended for Health Information Extraction Tool
'Copyright (C) 2020, Jacques Raubenheimer
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the code which does the actual extraction and parsing of data                   '
' from the Google Trends Web service.                                                                  '
' These results are meant to be similar to those obtained from the                                     '
' https://trends.google.com website, although they still require the API key.                          '
' These requests are also counted against a separate quota.                                            '
' The module calls some procedures in common with mGoogleTrendsInfoExtraction.                         '
' Because the data are different, it does not employ the multiple sampling approach                    '
' used in for the Google Trends Extended for Health data in mGoogleTrendsInfoExtraction.               '
' The calls use in this module (and specified on Sheet11--the Google Trends Web worksheet)             '
'  are listed on https://developers.google.com/apis-explorer/#p/trends/v1beta/                         '
' The full list is:                                                                                    '
' trends.getGraph                                                                                      '
'   Returns a Graph of search volume per time points, normalized. For better insights,                 '
'    one could provide restrictions for time range, geographic region, etc.                            '
' trends.getGraphAverages                                                                              '
'   Returns the averages of normalized search volume for the given terms.                              '
'    For better insights, one could provide restrictions for time range, geographic region, etc.       '
' trends.getRisingQueries                                                                              '
'   Get a list of rising queries that were searched along with the requested term,                     '
'    under the given restrictions.                                                                     '
' trends.getRisingTopics                                                                               '
'   Get a list of rising topics that were searched along with the requested term,                      '
'    under the given restrictions.                                                                     '
' trends.getTopQueries                                                                                 '
'   Get a list of top queries that were searched along with the requested term,                        '
'    under the given restrictions.                                                                     '
' trends.getTopTopics                                                                                  '
'   Get a list of top topics that were searched along with the requested term,                         '
'    under the given restrictions.                                                                     '
' trends.regions.list                                                                                  '
'   This would be the data behind the map seen in Regional Interest                                    '
'    in http://www.google.com/trends/explore                                                           '
''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''''    ''
' trends.getTimelinesForHealth                                                                         '
'   For health research only, fetches a graph of search volumes per time within a set of restrictions. '
'   Note the data is sampled and Google can't guarantee the accuracy of the numbers.                   '
' *This last one is the Google Trends Extended for Health API call, which is handled from              '
'  Sheet3 (Google Trends Extended Health) and mGoogleTrendsInfoExtraction.                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub GetGTWeb()
'This is the main procedure that is launched when the user clicks the Extract Data button on the
' Google Trends Web worksheet (Sheet11).
'It first does a complete check on each component needed for the extraction string as specified on the input worksheet,
' and then launches the procedure that sends the request to Google and builds the workbook with the returned data.
        
    Dim TimeProcessStart As Date            'Used to time how long the extraction process takes
    TimeProcessStart = Now
    
    Call StartSmoothly
    
    'Do an error check to ensure that all errors are resolved before beginning
    If Not fDoAllErrorCheckingForWeb Then EndGracefully

    Dim i As Integer                        'General purpose counter
    Dim sGTWebQueryString As String         'String that contains a complete query URL which will be passed to Google
    Dim sResult As String                   'These three variables are used to parse the JSON result returned by Google
    
                                            'These two variables store arrays of 1 to 3 elements (depending on the GTWeb function used)
                                            ' which provide the text to search for when parsing the JSON string.
                                            ' They are set in BuildArraysForWebParsing
    Dim sTopLevels() As String              'sTopLevels is used when the Graph function is called
    Dim sLabels() As String                 'sLabels is used when all other functions are called
    
    Dim vResult() As Variant                'An array storing the parsed data returned from the HTTP request
    Dim vTransposeResult() As Variant           'Created this 2020-09 to use in the TransposeArray output
    Dim wbkNew As Workbook                  'Stores the data extracted from Google Trends
    Dim wksQuerySpec As Worksheet           'Worksheet in wbkNew that stores the complete query specifications for auditing purposes
    Dim wksGTWebResults As Worksheet        'Worksheet in wbkNew where the actual data are written
    Dim rng As Range                        'General range variable
    Dim rng2 As Range
    
    'Variables used in creating a chart of the output
    Dim ch As Chart
    Dim rngX As Range       'Both of these are used in building the chart of the data plots
    Dim rngY As Range       'Both of these are used in building the chart of the data plots
    Dim rngHead As Range
    Dim sChartTitle As String
    Dim sValAxisTitle As String
    Dim sCatAxisTitle  As String
    
    'Build the URL to send to Google
    sGTWebQueryString = BuildGTWebQueryString
''    Debug.Print sGTWebQueryString
    
    'Define the labels to parse the string with, based on the Function called
    BuildArraysForWebParsing sFunction:=fvReturnNameValue("API_TermForMethod") _
                            , sArrTop:=sTopLevels _
                            , sArrLab:=sLabels
    
    'Pass the URL to Google and get the result back:
    sResult = fGetData(sGTWebQueryString)
    
    'For testing, write the returned string to a text file
    'At the moment, I am not testing whether fWriteToFile returns False as an indication of errors
    If Not bProductionVersion Then _
        fWriteToFile sFile:=Replace(fvReturnNameValue(sName:="WDataTarget", bCheckForActiveWorkbook:=False), ".xlsx", ".json", , , vbTextCompare) _
        , sFileContents:=sResult
    
    'Now parse the string
    '1. When values are passed, convert them to actual values
    '2. When multiple terms are passed to GetGraph,
    '   read each term's values to another element of vResults,
    '   which changes how vResult should be dimensioned.
    '3. When an error is returned, parse out the error information [Added 2020-09-30]
    '   Test below for an error and reset the parsing parameters accordingly
    If bInvalidArgumentError Or bOtherHTTPError Then _
    BuildArraysForWebParsing sFunction:="error" _
                            , sArrTop:=sTopLevels _
                            , sArrLab:=sLabels
    
    ParseGTWebResult V:=vResult _
                    , sArrTop:=sTopLevels _
                    , sArrLab:=sLabels _
                    , sJSON:=sResult
    
    'Add the new workbook to which the extracted data will be written
    Set wbkNew = Application.Workbooks.Add
    With wbkNew
        'Add a worksheet that will store the parsed data
        Set wksGTWebResults = .Sheets(1)
        With wksGTWebResults
            .Name = "Google Trends Web data"
            'Write the results to the worksheet
            Set rng = .Cells(1, 1)
            'rng.Resize(UBound(vResult, 1), UBound(vResult, 2)).value = Application.Transpose(vResult)
            '2020-09 Replace Application.Transpose with Chip Pearson's TransposeArray
            TransposeArray vResult, vTransposeResult
            rng.Resize(UBound(vResult, 2), UBound(vResult, 1)).Value2 = vTransposeResult
'            rng.Resize(UBound(vResult, 2), UBound(vResult, 1)).Value = vResult
            
            'rng.resize does not actually change the address of the range, so it must be reset
            Set rng = .UsedRange
            If fvReturnNameValue("API_TermForMethod", bCheckForActiveWorkbook:=False) = "graph" Then _
                Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1)).NumberFormat = "yyyy-mm-dd"
            
            'Copy the date range to the worksheet
            ' This use useful here, especially for Rising/Top Queries/Topics
            ' (i.e., the functions which do not generate charts)
            Set rng2 = .Range(wksGTWebResults.Cells(1, rng.Columns.Count + 2), .Cells(3, rng.Columns.Count + 3))
            CopyRange rngTarget:=rng2 _
                    , rngSource:=Sheet11.Range("CompleteDateRangeSpecificationWeb")
            rng2.Cells(1, 1).Style = "Heading 2"
            rng2.Columns.AutoFit
        End With
        
        'Format the results
        With rng
            .Rows.RowHeight = 15
            .Rows(1).Font.Bold = True
            .WrapText = False
            .Columns.AutoFit
        End With
        'Check that no columns are too wide after the autofitting
        DialBackColumnWidths wks:=wksGTWebResults
        
        'If the graph or region functions were called, create a chart showing the data
        
        'Build the chart title
        sChartTitle = "Google Trends Web data: " & fvReturnNameValue("WFunction", False)
        sChartTitle = sChartTitle & vbLf & "Geographic location: "
        If fvReturnNameValue("WGeographicLevel", False) = "Worldwide" Then
            sChartTitle = sChartTitle & "Worldwide"
        Else
            'Add the country
            sChartTitle = sChartTitle & fvReturnNameValue("WCountry", False)
            'If region, then add the region too
            If fvReturnNameValue("WGeographicLevel", False) = "Region" Then _
                sChartTitle = sChartTitle & "-" & fvReturnNameValue("WRegion", False)
        End If

        'Add the time frame
        sChartTitle = sChartTitle & vbLf _
            & "Time range: " _
            & Format(IIf(Len(fvReturnNameValue("WStartDate", False)) = 0, #1/1/2004#, fvReturnNameValue("WStartDate", False)), "dd mmm yyyy") & " to " _
            & Format(IIf(Len(fvReturnNameValue("WEndDate", False)) = 0, Now, fvReturnNameValue("WEndDate", False)), "dd mmm yyyy")

        'Add the terms
        sChartTitle = sChartTitle & vbLf & "Terms: "
        For i = 1 To 5
            If Len(fvReturnNameValue("WSearchTerm0" & i, False)) > 0 Then _
                sChartTitle = sChartTitle & IIf(i > 1, "||", vbNullString) _
                    & fvReturnNameValue("WSearchTerm0" & i, False)
        Next i
        'Check that the chart title is less than the limit of 255 characters
        If Len(sChartTitle) > 255 Then _
        sChartTitle = Left(sChartTitle, 252) & Chr(133)
        
        'Build the line chart for the graph function
        If fvReturnNameValue("API_TermForMethod", bCheckForActiveWorkbook:=False) = "graph" Then
            
            'Build the axis titles
            sValAxisTitle = "Relative probability of search occurrence scaled to 100"
'            sCatAxisTitle = "Date range"
            
            Set rngX = Intersect(rng.Columns(1), Range(rng.Rows(2), rng.Rows(rng.Rows.Count)))
            
            'Add the chart
            Set ch = fAddChart(ct:=xlLine _
                            , wb:=wbkNew _
                             , sSub:="Plot of Google Trends Web data" _
                             , RngXVal:=rngX _
                             , sTitle:=sChartTitle _
                             , bAddLegend:=(UBound(vResult, 1) > 2) _
                             , bScaleValuesTo100:=True _
                             , sValAxisTitle:=sValAxisTitle _
                             , bAddDataLabels:=False)
            'Add the series to the chart
            wksGTWebResults.Activate
            For i = 2 To UBound(vResult, 1)
                Set rngY = Intersect(rng.Columns(i), Range(rng.Rows(2), rng.Rows(rng.Rows.Count)))
                Set rngHead = Intersect(rng.Rows(1), rng.Columns(i))
                AddSeries ch:=ch _
                        , iSeriesNo:=i - 1 _
                        , rngNameAddress:=rngHead _
                        , rngData:=rngY _
                        , RngXVal:=rngX _
                        , bSetColourToBlack:=(UBound(vResult, 1) <= 2)
            Next i
        'Build the bar chart for the List of regions
        ElseIf fvReturnNameValue("API_TermForMethod", bCheckForActiveWorkbook:=False) = "regions" Then
            
            'Build the axis titles
            If fvReturnNameValue("WGeographicLevel", False) = "Worldwide" Then
                sCatAxisTitle = "World regions"
            ElseIf fvReturnNameValue("WGeographicLevel", False) = "Country" Then
                sCatAxisTitle = fvReturnNameValue("WCountry", False) & ": Regions"
            ElseIf fvReturnNameValue("WGeographicLevel", False) = "Region" Then
                sCatAxisTitle = fvReturnNameValue("WCountry", False) & " (" & fvReturnNameValue("WRegion", False) & "): Cities"
            End If
            sValAxisTitle = "Relative probability of search occurrence scaled to 100"
            
            Set rngX = Intersect(rng.Columns(2), Range(rng.Rows(2), rng.Rows(rng.Rows.Count)))
            
            'Add the chart
            Set ch = fAddChart(ct:=xlBarClustered _
                             , wb:=wbkNew _
                             , sSub:="Plot of Region values" _
                             , RngXVal:=rngX _
                             , sTitle:=sChartTitle _
                             , bAddLegend:=False _
                             , bScaleValuesTo100:=True _
                             , sValAxisTitle:=sValAxisTitle _
                             , sCatAxisTitle:=sCatAxisTitle _
                             , bAddDataLabels:=True)
            
            'Add the series to the chart
            wksGTWebResults.Activate
            i = 3
            Set rngY = Intersect(rng.Columns(i), Range(rng.Rows(2), rng.Rows(rng.Rows.Count)))
            Set rngHead = Intersect(rng.Rows(1), rng.Columns(i))
            AddSeries ch:=ch _
                    , iSeriesNo:=i - 2 _
                    , rngNameAddress:=rngHead _
                    , rngData:=rngY _
                    , RngXVal:=rngX _
                    , bSetColourToBlack:=False
        
        'Possible to-do item:
        'Currently, no charts are generated for:
        '- topTopics
        '- topQueries
        '- risingTopics
        '- risingQueries
        '- graphAverages

        End If
        
        'Add the query specification auditing sheet
        Set wksQuerySpec = .Sheets.Add(before:=.Sheets(1))
        
        'Collect all the query terms into the array vQueryList.
        'This is different to BuildQueryList which returns a single string will all the query terms preceded by "&term="
        ' Check through the whole range, just to be sure
        'Trim away leading and trailing spaces as well
        Dim vTerms() As Variant
        Dim sTmp As String
        For i = 1 To 5
            sTmp = Trim$(fvReturnNameValue(sName:="WSearchTerm" & Format(i, "00"), bCheckForActiveWorkbook:=False))
            If Len(sTmp) > 0 Then
                ReDim Preserve vTerms(1 To i)
                vTerms(i) = sTmp
            End If
        Next i
        sTmp = vbNullString
        
        'Store the regions list
        Dim vRegions() As Variant
        If fvReturnNameValue("WGeographicLevel") <> "Region" Then
            ReDim vRegions(1 To 1)
            vRegions(1) = vbNullString
        Else
            If Len(fvReturnNameValue("WRegion")) = 0 Then
                vRegions = Range("WGeoAllRegionsForCountry")
            Else
                ReDim vRegions(1 To 1)
                vRegions(1) = fvReturnNameValue("WRegion")
            End If
        End If
        
        'Summarise the query specifications
        SetUpQuerySpecSheetWeb wksQuerySpec:=wksQuerySpec _
                             , vQueryList:=vTerms _
                             , vRegions:=vRegions _
                             , vURL:=sGTWebQueryString _
                             , StartTime:=TimeProcessStart
        
        'Save the workbook
        'wbkNew.SaveAs fvReturnNameValue(sName:="WDataTarget", bCheckForActiveWorkbook:=False)
        SaveWithErrorHandling iSaveOrSaveAs:=2 _
                            , wbk:=wbkNew _
                            , sFilePath:=fvReturnNameValue(sName:="WDataTarget", bCheckForActiveWorkbook:=False) _
                            , bAddToMRU:=False

    End With
    
    'Update the Target file message to show that this file now exists
    With Sheet11
        .Range("WQueriesThisSession").Value = fvReturnNameValue("WQueriesThisSession") + fvReturnNameValue("WQueriesNeeded")
        .Range("WDataTarget").Value = .Range("WDataTarget").Value
        .Range("WTargetFileMessage1").Calculate
    End With
    
    If bShowCompletionMsgBoxes Then
        MsgBox "The Google Trends Web data extraction has completed successfully." _
            , vbInformation + vbOKOnly, "Google Trends Web"
        
        wbkNew.Activate
        
        'Turn on Cut/CopyPaste for the new workbook
        TurnOnPaste
        
        EndGracefully
    Else
        wbkNew.Close SaveChanges:=True
    End If
    
End Sub

Private Sub ParseGTWebResult(ByRef V() As Variant _
                           , ByRef sArrTop() As String _
                           , ByRef sArrLab() As String _
                           , ByRef sJSON As String)
'Uses the labelling arrays sArrTop and sArrLab to parse the JSON string (passed as sJSON)
' and extract all the values to the array v
'sArrTop and sArrLab are two arrays containing the JSON keys appearing in the string (as defined by the API), which can be used to grab the contents
'v will always be two-dimensional.
' The number of elements in the first dimension will be determined by the function called
' and the number of elements in the second dimension will be determined by the returned string
' One element is added to the second dimension to contain the labels, and for graph only, to the first dimension to contain the dates
'XX (this will be incrementally increased using reDim Preserve)

    Dim sSearchString1 As String            'Element to search for in JSON string
    Dim sSearchString2 As String            'Element to search for in JSON string
    Dim lPISstart As Long                   'A counter to determine the starting point for the data being extracted from the JSON string
    Dim lPISend As Long                     'A counter to determine the end point for the data being extracted from the JSON string
    Dim iCounter As Integer                 'A counter that loops through elements of the JSON string
    Dim aTermPoints() As Long               'Stores the points in the JSON string at which the phrase "term: " is found
                                            ' so that various terms in the query can be parsed out
    Dim iTermCounter As Integer             'Count which term is being extracted
    Dim lPosChecker As Long                 'Used to check whether a term's values still exist within the remaining string section to parse
    Dim iSkipFirstColWithGraph As Integer   'With the graph function, one column must be skipped on the first iteration
                                            'The two variables below are needed because the JSON string returned by Google
                                            ' for some functions contains the values for each term in succession, so in essence,
                                            ' every row for the first term must be written in the first column,
                                            ' then every row for the second term must be written in the second column, etc.
    Dim iDestinationColumn As Integer       'Indicates which column in the destination array to write the value to
    Dim iDataPointCounter As Integer        'Counts which row in the destination array the value must be written to
    
    Dim vTerms As Variant               'Array to store all terms for which search volumes were retrieved
    Dim vDates As Variant               'Array to store all dates in returned sampling
    Dim vValues As Variant              'Array of all values for any particular term, for all dates in the sampling
    Dim dicResult As Object         'Added 2020-08-07 to parse JSON data. To be late-bound below
    
    '2020-08-07
    'Parse the JSON string using Daniel Ferry's JSON parser
    Set dicResult = CreateObject("Scripting.Dictionary")
    Set dicResult = ParseJSON(sJSON)
    
    'First determine which Function was called from Google, based on the first element of sArrTop
    'If it was graph, then dimension aTermPoints to locate all terms in the string
    'For GetGraph, one column for the dates, and one for each term
    'For all others, the number of labels

    If sArrTop(1) = "error" Then    'Report the error [added 2020-09-30]
        'Transform the labels into error description terms for the JSON parsing
        'These are stored in vTerms. It has to be zero-bound for the GetFilteredTable procedure
        ReDim vTerms(0 To UBound(sArrLab) - 1)
        ReDim V(1 To UBound(sArrLab), 1 To 1)
        Dim sTmpError As String
        
        For iTermCounter = 0 To UBound(sArrLab) - 1
            If iTermCounter <> 0 And iTermCounter <> 4 Then sTmpError = "errors(0)." Else sTmpError = vbNullString
            vTerms(iTermCounter) = "*." & sTmpError & sArrLab(iTermCounter + 1) & "*"
            'Put the labels into V for output
            V(iTermCounter + 1, 1) = sArrLab(iTermCounter + 1)
        Next iTermCounter
        'Get all the terms as a table (two-dimensional array)
        vValues = GetFilteredTable(dicResult, vTerms)
        'Put the values into V for output
        'First redimension v to be 1 bigger than the values (labels are already in first element)
        'V is an array that must be transposed. The first array dimension is rows, the second is columns.
        'When V is written to the worksheet, Application.transpose is used to reverse that.
        'The data from GetFilteredTable stored in vValues is, however, also 'transposed'--
        ' the first dimension grows as the number of terms does.
        ' So the first dimension of vValues must be stored in the second dimension of V, and vice versa.
        ReDim Preserve V(LBound(V) To UBound(V), LBound(vValues, 1) To UBound(vValues, 1) + 1)
        For iTermCounter = LBound(vValues, 1) To UBound(vValues, 1)
            For iCounter = LBound(vValues, 2) To UBound(vValues, 2)
                V(iCounter, iTermCounter + 1) = vValues(iTermCounter, iCounter)
            Next iCounter
        Next iTermCounter
    
    
    
    
    ElseIf sArrTop(1) = "lines" Then    'API call: graph
        'Get the list of terms from the JSON string
        vTerms = GetFilteredValues(dicResult, "*.term*")
        'Get the dates and values for each term from the JSON string
        vDates = GetFilteredValues(dicResult, "*.lines(0).*date*")
        'Dimension v to contain the data
        'First dimension is one more than the number of terms (first column contains dates)
        'Second dimension is one more than the number of values (first row contains 'titles': 'date' and each term
        ReDim V(1 To UBound(vTerms) + 1, 1 To UBound(vDates) + 1) 'Add one because the dates and labels are added in the first element of the two dimensions, respectively
        'Add the dates to the first dimension
        'First add the label
        V(1, 1) = "date"
        For iTermCounter = LBound(vDates) To UBound(vDates)
            V(1, iTermCounter + 1) = DateValue(vDates(iTermCounter))
            'Error check the date calculation
            'Stricly speaking, all errors in the process should have been intercepted in fGetData (fTestHTTPDataForErrors)
            If Err.Number <> 0 Then
                Err.Clear
        '       Debug.Print Mid$(sResult, lPISstart, lPISend - lPISstart)
                MsgBox "An error occured when attempting to interpret the date " _
                    & vDates(iTermCounter) _
                    & " found for value [" & iTermCounter & "]" _
                    & " from the JSON string returned by Google Trends." _
                    & vbCrLf & "Please check the query specification and retry the process." _
                    , vbCritical + vbOKOnly, "Invalid date found"
                
                'Do not end completely, but exit the function, so that existing data can be reported
                Exit Sub
            End If
        Next iTermCounter
        'Add the terms and their values for all dates
        For iTermCounter = LBound(vTerms) To UBound(vTerms)
            'Add the terms to the first row
            V(iTermCounter + 1, 1) = vTerms(iTermCounter) 'iTermCounter+1 to make place for the date column
            'Read the values and add them to the remainder of the column array dimension
            vValues = GetFilteredValues(dicResult, "*.lines(" & iTermCounter - 1 & ").*value*")    'iTermCounter-1 because the dictionary is zero-based
            For iCounter = LBound(vValues) To UBound(vValues)
                V(iTermCounter + 1, iCounter + 1) = vValues(iCounter)
            Next iCounter
        Next iTermCounter
    Else
        'Transform the labels into match terms for the JSON parsing
        'These are stored in vTerms. It has to be zero-bound for the GetFilteredTable procedure
        ReDim vTerms(0 To UBound(sArrLab) - 1)
        ReDim V(1 To UBound(sArrLab), 1 To 1)
        For iTermCounter = 0 To UBound(sArrLab) - 1
            vTerms(iTermCounter) = "*." & sArrLab(iTermCounter + 1) & "*"
            'Put the labels into V for output
            V(iTermCounter + 1, 1) = sArrLab(iTermCounter + 1)
        Next iTermCounter
        'Get all the terms as a table (two-dimensional array)
        vValues = GetFilteredTable(dicResult, vTerms)
        'Put the values into V for output
        'First redimension v to be 1 bigger than the values (labels are already in first element)
        'V is an array that must be transposed. The first array dimension is rows, the second is columns.
        'When V is written to the worksheet, Application.transpose is used to reverse that.
        'The data from GetFilteredTable stored in vValues is, however, also 'transposed'--
        ' the first dimension grows as the number of terms does.
        ' So the first dimension of vValues must be stored in the second dimension of V, and vice versa.
        ReDim Preserve V(LBound(V) To UBound(V), LBound(vValues, 1) To UBound(vValues, 1) + 1)
        For iTermCounter = LBound(vValues, 1) To UBound(vValues, 1)
            For iCounter = LBound(vValues, 2) To UBound(vValues, 2)
                V(iCounter, iTermCounter + 1) = vValues(iTermCounter, iCounter)
            Next iCounter
        Next iTermCounter
    End If

'--- This code is commented out, as it is obviated by the use of Daniel Ferry's JSON parser ---'
''''    GoTo SkipOldGTWeb
''''    If sArrTop(1) = "lines" Then
''''        SetUpArrayGraphFunction aTermPoints:=aTermPoints _
''''                              , V:=V _
''''                              , sArrTop:=sArrTop _
''''                              , sJSON:=sJSON
''''    Else    'sArrTop(1) = "averages",sArrTop(1) = "regions",sArrTop(1) = "item"
''''        SetUpArraysNonGraphFunctions aTermPoints:=aTermPoints _
''''                                    , V:=V _
''''                                    , sArrLab:=sArrLab _
''''                                    , sJSON:=sJSON
''''    End If
''''
''''    'Parse the string element-by-element
''''    'For Graph, UBound(aTermPoints) will be the number of terms searched + 1
''''    'For all other functions, UBound(aTermPoints) will be UBound(sArrLab) + 1
''''    For iTermCounter = LBound(aTermPoints) To UBound(aTermPoints) - 1   'The last value in aTermPoints is the end (length) of the string, so that is excluded
''''
''''        If iTermCounter = 1 Then iSkipFirstColWithGraph = 0 Else iSkipFirstColWithGraph = 1
''''
''''        iDataPointCounter = 1   'Set to 1, not 0, so that the header row is preserved
''''
''''        'Set the starting point to that position in the string where the term is encountered
''''        lPISstart = aTermPoints(iTermCounter)
''''
''''        'Loop through the string, searching for the labels
''''        Do 'While lPISstart <> 0
''''            'Check that we have not moved into the next term's values
''''            lPosChecker = InStr(lPISstart, sJSON, sQuote & sArrLab(1) & sQuote & ": " _
''''                & IIf(sArrLab(1) = "value" Or sArrLab(1) = "isBreakout", vbNullString, sQuote), vbTextCompare)
''''            If lPosChecker > aTermPoints(iTermCounter + 1) Or lPosChecker = 0 Then
''''                lPISstart = 0
''''            Else
''''                iDataPointCounter = iDataPointCounter + 1
''''                If iTermCounter = 1 Then ReDim Preserve V(LBound(V, 1) To UBound(V, 1), 1 To iDataPointCounter)
''''
''''                For iCounter = LBound(sArrLab) + iSkipFirstColWithGraph To UBound(sArrLab)
''''                    If iTermCounter = 1 Then
''''                        iDestinationColumn = iCounter
''''                    Else
''''                        iDestinationColumn = iTermCounter + 1
''''                    End If
''''
''''                    'Extract the data element to the array
''''                    ExtractOneElement lStart:=lPISstart _
''''                                    , lEnd:=lPISend _
''''                                    , sJSON:=sJSON _
''''                                    , sLabs:=sArrLab _
''''                                    , iLabNo:=iCounter _
''''                                    , V:=V _
''''                                    , iRow:=iDataPointCounter _
''''                                    , iCol:=iDestinationColumn
''''
''''                    If lPISstart = 0 Then Exit For  'If the first label has not been found, cease searching for any other labels
''''                Next iCounter   '= LBound(sArrLab) To UBound(sArrLab)
''''            End If
''''
''''        Loop While lPISstart <> 0
''''
''''    Next iTermCounter   '= LBound(aTermPoints) To UBound(aTermPoints)
''''SkipOldGTWeb:
End Sub

'--- This code (all the procedures below) is commented out, as it is obviated by the use of Daniel Ferry's JSON parser ---'
''''    'First determine which Function was called from Google, based on the first element of sArrTop
''''    'If it was graph, then dimension aTermPoints to locate all terms in the string
''''    'For GetGraph, one column for the dates, and one for each term
''''    'For all others, the number of labels
''''
''''Sub SetUpArraysNonGraphFunctions(ByRef aTermPoints() As Long _
''''                               , ByRef V() As Variant _
''''                               , ByRef sArrLab() As String _
''''                               , ByRef sJSON As String)
'''''This procedure dimensions two arrays [which are passed ByRef] for parsing the values from the returned JSON string.
'''''1. aTermPoints() will be used to look through the JSON string and store the starting points
'''''    for each of the 1 or more terms contained in the JSON string.
'''''    These points will then tell the parser where to start and stop extracting values for each term.
'''''    For non-graph functions, there can be only on term,
'''''    so aTermPoints() is simply set to the start (1) and the end (Len) of the JSON string.
'''''2. v() is the array which will contain the parsed values.
'''''    Because only the last dimension can be resized with ReDim Preserve, it is actually a transposed array,
'''''    with rows (first dimension) for each "field" and columns (second dimension) for each new value.
'''''    It contains a row (first dimension) for every label (in all functions except Graph), in which to store the values.
'''''    There are between 2 and 4 labels, depending on which function was called.
'''''    The labels are contained in sArrLab() and are set in BuildArraysForWebParsing.
'''''    v() will eventually contain one more column than the number of returned values, as the first column is used to
'''''    store the labels. When v() is transposed back into Excel, this first column will form the row containing the column headers.
''''
''''    'Set the starting and ending points for the search in the JSON string
''''    ReDim aTermPoints(1 To 2)
''''    aTermPoints(1) = 1          'Since we are extracting only one column, simply start at the beginning...
''''    aTermPoints(2) = Len(sJSON) '... and add the end of the JSON string as the last point for aTermPoints
''''
''''    'Set the v array to contain the right number of columns
''''    ReDim V(LBound(sArrLab) To UBound(sArrLab), 1 To 1)
''''
''''    'Now add the labels to the first element
''''    Dim i As Integer
''''    For i = LBound(sArrLab) To UBound(sArrLab)
''''        V(i, 1) = sArrLab(i)
''''    Next i
''''
''''End Sub
''''
''''Sub SetUpArrayGraphFunction(ByRef aTermPoints() As Long _
''''                          , ByRef V() As Variant _
''''                          , ByRef sArrTop() As String _
''''                          , ByRef sJSON As String)
'''''This procedure dimensions two arrays [which are passed ByRef] for parsing the values from the returned JSON string.
'''''1. aTermPoints() will be used to look through the JSON string and store the starting points
'''''    for each of the 1 or more terms contained in the JSON string.
'''''    These points will then tell the parser where to start and stop extracting values for each term.
'''''    For the graph function, there may be a single or multiple terms,
'''''    so the values for aTermPoints() are set by looping through the JSON string and
'''''    finding each occurrence of the 2nd-level sArrTop value in context ["term": "].
'''''2. v() is the array which will contain the parsed values.
'''''    Because only the last dimension can be resized with ReDim Preserve, it is actually a transposed array,
'''''    with rows (first dimension) for each "field" and columns (second dimension) for each new value.
'''''    It contains a row (first dimension) for the dates, as well as rows for the values of each term.
'''''    The term identifiers are contained in sArrTop() and are set in BuildArraysForWebParsing.
'''''    It will eventually contain one more column than the number of returned values, as the first column is used to
'''''    store the labels. When v() is transposed back into Excel, this first column will form the row containing the column headers.
''''
''''    Dim lStart As Long                  'Moving starting point, for working through the JSON string
''''    Dim iTermCounter As Long            'Counts how many terms are idendified in the JSON string
''''    Dim sSearchStringAtStart As String  'String that idenfies the start of a set of term values in the JSON string
''''    Dim sSearchStringAtEnd As String    'String that idenfies the end of a set of term values in the JSON string
''''
''''    'sArrTop(2)="term" for the graph JSON string. This will mark the start of the data for each successive term queried
''''    sSearchStringAtStart = sQuote & sArrTop(2) & sQuote & ": " & sQuote
''''    sSearchStringAtEnd = sQuote & ","
''''    'The third level ("points") can just be ignored
''''    'sSearchStringAtEnd = sQuote & "," & sQuote & sArrTop(3) & sQuote & ": [{"
''''    'sSearchStringAtEnd = sQuote & "," & sQuote & "points" & sQuote & ": [{"
''''
''''    lStart = 1
''''    iTermCounter = 0
''''
''''    'Loop through the string and find every occurrence of "term" which will indicate the number of terms queried and returned
''''    Do
''''        lStart = InStr(lStart, sJSON, sSearchStringAtStart, vbTextCompare)
''''        If lStart > 0 Then
''''            'Move the start forward, so that the next item can be found on the next iteration
''''            ' and this also makes the eventual string search more efficient, and allows us to extract the term as a label below
''''            lStart = lStart + Len(sSearchStringAtStart)
''''            iTermCounter = iTermCounter + 1
''''            ReDim Preserve aTermPoints(1 To iTermCounter)
''''            aTermPoints(iTermCounter) = lStart
''''        End If
''''    Loop While lStart <> 0
''''
''''    'Add the end of the JSON string as the last point for aTermPoints
''''    ReDim Preserve aTermPoints(1 To UBound(aTermPoints) + 1)
''''    aTermPoints(UBound(aTermPoints)) = Len(sJSON)
''''
''''    'Redimension the v() array according to the count of the number of terms (UBound(aTermPoints)-1) which has just been calculated
''''    'However, add an extra column to store the date values, so ReDim v() to UBound(aTermPoints)-1+1=UBound(aTermPoints)
''''    ReDim V(1 To UBound(aTermPoints), 1 To 1)
''''
''''    'Now add the term labels
''''    V(1, 1) = "Date"
''''    For iTermCounter = 1 To UBound(aTermPoints) - 1
''''
''''        'Also Remove quote Escape characters "\\\"
''''        V(iTermCounter + 1, 1) = Replace(Mid(sJSON, aTermPoints(iTermCounter) _
''''                                   , InStr(aTermPoints(iTermCounter), sJSON, sSearchStringAtEnd, vbTextCompare) _
''''                                   - aTermPoints(iTermCounter)) _
''''                                 , "\\\" & sQuote, sQuote, , , vbTextCompare)
''''    Next iTermCounter
''''
''''End Sub
''''
''''Sub ExtractOneElement(ByRef lStart As Long _
''''                    , ByRef lEnd As Long _
''''                    , ByRef sJSON As String _
''''                    , ByRef sLabs() As String _
''''                    , ByVal iLabNo As Integer _
''''                    , ByRef V() As Variant _
''''                    , ByVal iRow As Integer _
''''                    , Optional ByVal iCol As Integer)
''''
'''''This procedure uses the defined start (lStart) and end (lEnd) points and searches the JSON string
''''' for a value to extract, using the supplied label from the array (iLabNo from sLabs()),
''''' and then extracts that value and stores it in the defined row (iRow) and column (iCol) '
''''' of the array of parsed values (v())
''''
''''    'iCol is used to parcel out graph values for multiple term queries to different columns
''''    'if the query is not a graph query, then the value is simply extracted to the same column as the label in question
''''    If iCol = 0 Then iCol = iLabNo
''''
''''    Dim sSearchStringPreceding As String      'Element to search for in JSON string
''''    Dim sSearchStringSucceeding As String     'Element to search for in JSON string
''''    Dim bIsString As Boolean         'Test whether the element we are extracting is a string (enclosed in quotes)
''''                                     ' or whether it is a value (number for 'value' and boolean for isBreakout)
''''    'Set the search string type
''''    If sLabs(iLabNo) = "value" Or sLabs(iLabNo) = "isBreakout" Then bIsString = False Else bIsString = True
''''
''''    sSearchStringPreceding = sQuote & sLabs(iLabNo) & sQuote & ": "
''''    'If it is the last term being extracted, search for a brace, otherwise, search for a comma
''''    sSearchStringSucceeding = IIf(iLabNo = UBound(sLabs), "}", ",")
''''    'if the term is a string term, add the quotes which must be included in the search
''''    If bIsString Then
''''        sSearchStringPreceding = sSearchStringPreceding & sQuote
''''        sSearchStringSucceeding = sQuote & sSearchStringSucceeding
''''    End If
''''
''''    'Find the label
''''    lStart = InStr(lStart, sJSON, sSearchStringPreceding, vbTextCompare)
''''    'If the label is found, set the limits of the value to be extracted
''''    ' by extending the starting position, and defining the end position
''''    If lStart <> 0 Then
''''        lStart = lStart + Len(sSearchStringPreceding)
''''        lEnd = InStr(lStart, sJSON, sSearchStringSucceeding, vbTextCompare)
''''
''''        'Store the value
''''        If sLabs(iLabNo) = "value" Then     'Store a numeric value
''''            V(iCol, iRow) = CDbl(Mid(sJSON, lStart, lEnd - lStart))
''''        ElseIf sLabs(iLabNo) = "date" Then  'Store a date value
''''            'v(iCol, iRow) = fParseDate(sDate:=Mid(sJSON, lStart, lEnd - lStart))
''''            V(iCol, iRow) = Format(CDate(Mid(sJSON, lStart, lEnd - lStart)), "yyyy-mm-dd")
''''        Else                                'Store string values
''''            V(iCol, iRow) = Mid(sJSON, lStart, lEnd - lStart)
''''        End If
''''
''''        'Move the start position forward
''''        lStart = lEnd + Len(sSearchStringSucceeding)
''''
''''    End If
''''End Sub
''''
''''Function fParseDate(sDate As String) As Date
'''''No longer used.
''''' Was used in ExtractOneElement to successfully extract date values
''''    Dim y As Integer
''''    Dim m As Integer
''''    Dim d As Integer
''''    Dim sSeparator As String * 1
''''    Dim iSep1 As Integer
''''
''''    If fFindDateSeparator(sDate) = vbNullString Then
''''        MsgBox "Cannot find a valid date separator in the data returned from Google." _
''''            & vbCrLf & "The date string is: " & sDate, vbCritical + vbOKOnly, "No date separator"
''''        EndGracefully
''''    End If
''''    sSeparator = fFindDateSeparator(sDate)
''''    y = CInt(Left(sDate, 4))
''''    iSep1 = InStr(1, sDate, sSeparator, vbTextCompare)
''''    m = CInt(Mid(sDate, iSep1 + 1, InStr(iSep1 + 1, sDate, sSeparator, vbTextCompare) - iSep1 - 1))
''''    d = CInt(Right(sDate, 2))
''''    On Error Resume Next
''''    fParseDate = DateSerial(y, m, d)
''''    'Debug.Print "y:" & y & " m:" & m & " d:" & d & " " & fParseDate
''''    If Err.Number <> 0 Then
''''        MsgBox "A date string (" & sDate & ") in the data returned from Google cannot be parsed into a valid date." _
''''            , vbCritical + vbOKOnly, "Invalid date"
''''        EndGracefully
''''    End If
''''End Function
''''
''''Function fFindDateSeparator(sDate As String) As String
'''''No longer used. Called from fParseDate.
''''    If Len(Replace(sDate, "/", vbNullString, , , vbTextCompare)) = Len(sDate) - 2 Then
''''        fFindDateSeparator = "/"
''''    ElseIf Len(Replace(sDate, "-", vbNullString, , , vbTextCompare)) = Len(sDate) - 2 Then
''''        fFindDateSeparator = "-"
''''    Else
''''        fFindDateSeparator = vbNullString
''''    End If
''''End Function

Private Sub BuildArraysForWebParsing(ByVal sFunction As String _
                           , ByRef sArrTop() As String _
                           , ByRef sArrLab() As String)
'Takes the function being called from Google as a string, and the builds two arrays (passed ByRef)
' which will be used for parsing the JSON string that Google returns.
' The string is different for each function, so these arrays will allow one procedure to parse any returned string

'Possible functions:
'Graph                  getGraph            graph
'Graph Averages         getGraphAverages    graphAverages
'List of Region Values  regions.list        regions
'Rising Queries         getRisingQueries    risingQueries
'Rising Topics          getRisingTopics     risingTopics
'Top Queries            getTopQueries       topQueries
'Top Topics             getTopTopics        topTopics
        
    'Create an array containing the top level labels. These mostly appear only once, and will actually mostly be ignored
    'Select Case LCase$(sFunction)
    Select Case sFunction
    Case "error"    'Added 2020-09-30
        ReDim sArrTop(1 To 1)
        sArrTop(1) = "error"
    Case "graph"
        ReDim sArrTop(1 To 3)
        sArrTop(1) = "lines"
        sArrTop(2) = "term"
        sArrTop(3) = "points"
    Case "graphAverages"
        ReDim sArrTop(1 To 1)
        sArrTop(1) = "averages"
    Case "regions"
        ReDim sArrTop(1 To 1)
        sArrTop(1) = "regions"
    Case Else   '        getTopTopics, getTopQueries, getRisingTopics, getRisingQueries
        ReDim sArrTop(1 To 1)
        sArrTop(1) = "item"
    End Select
    
    'Create an array containing the labels across which values will be retrieved
    'Select Case LCase$(sFunction)
    Select Case sFunction
    Case "error"    'Added 2020-09-30
        ReDim sArrLab(1 To 5)
        sArrLab(1) = "code"
        sArrLab(2) = "message"
        sArrLab(3) = "domain"
        sArrLab(4) = "reason"
        sArrLab(5) = "status"
    Case "regions"
        ReDim sArrLab(1 To 3)
        sArrLab(1) = "regionCode"
        sArrLab(2) = "regionName"
        sArrLab(3) = "value"
    Case "topTopics"
        ReDim sArrLab(1 To 3)
        sArrLab(1) = "title"
        sArrLab(2) = "mid"
        sArrLab(3) = "value"
    Case "topQueries"
        ReDim sArrLab(1 To 2)
        sArrLab(1) = "title"
        sArrLab(2) = "value"
    Case "risingTopics"
        ReDim sArrLab(1 To 4)
        sArrLab(1) = "title"
        sArrLab(2) = "mid"
        sArrLab(3) = "value"
        sArrLab(4) = "isBreakout"
    Case "risingQueries"
        ReDim sArrLab(1 To 3)
        sArrLab(1) = "title"
        sArrLab(2) = "value"
        sArrLab(3) = "isBreakout"
    Case "graphAverages"
        ReDim sArrLab(1 To 2)
        sArrLab(1) = "term"
        sArrLab(2) = "value"
    Case "graph"
        ReDim sArrLab(1 To 2)
        sArrLab(1) = "date"
        sArrLab(2) = "value"
    End Select


End Sub

Private Function BuildGTWebQueryString() As String
'This procedure reads the values for the GT Web query specification from the Google Trends Web worksheet (Sheet11)
' and builds a URL which can be sent to Google via the HTTP request.

    Dim sURL As String
    
    'Build the base of the URL
    sURL = sReqStrGTW_URL_Base & fvReturnNameValue("API_TermForMethod") & "?"
    
    'Add the terms
    sURL = sURL & BuildQueryList(Sheet11)
    
    'Build Start Date
    sURL = sURL & fBuildQueryComponent(sReqStrGTW_StartDate, Format(fvReturnNameValue("WStartDate"), "YYYY-MM"))
    
    'Build End date
    sURL = sURL & fBuildQueryComponent(sReqStrGTW_EndDate, Format(fvReturnNameValue("WEndDate"), "YYYY-MM"))
    
    'Build geographic restrictions
    Select Case fvReturnNameValue("WGeographicLevel")
    Case "Worldwide"
        'Do nothing
    Case "Country"
        sURL = sURL & fBuildQueryComponent(sReqStrGTW_GeoRestr, fvReturnNameValue("WDescriptor"))
    Case "Region"
        'Possible to-do item (see SetUpQuerySpecSheetWeb as well):
        'Modify this to do a multi-region search for GetGraph
        sURL = sURL & fBuildQueryComponent(sReqStrGTW_GeoRestr, fvReturnNameValue("WDescriptor"))
    End Select
    
    'Build property restrictions (Search domain)
    sURL = sURL & fBuildQueryComponent(sReqStrGTW_Prop, fvReturnNameValue("DomainForWebQuery"))
    
    'Build category restrictions
    Dim sTmpCat As String
    sTmpCat = fvReturnNameValue("WCategory")
    If Len(sTmpCat) > 0 Then _
        sURL = sURL & fBuildQueryComponent(sReqStrGTW_Cat, Trim(Split(sTmpCat, ":", , vbTextCompare)(1)))
    
    'Build field restrictions
    'Not currently used, as the Field restrictions limit what is returned
    ' (e.g., excluding the date, or excluding the value), and it is better to have everything returned
    'sURL = sURL & fBuildQueryComponent(sReqStrGTW_Fields, "averages" "lines(points/value,term)" etc.)
    
    'Possible to-do item:
    'Public Const sReqStrGTEH_TimelReso As String = "&timelineResolution="
    'Attempt to get other time resolutions than the default supplied by the Google Trends website
    
    'Add API key
    sURL = sURL & sReqStrKey & fGetValueFromFile(fvReturnNameValue("APIKey"))
    
    'Remove conflicting signifiers
    sURL = Replace(sURL, "?&", "?", , , vbTextCompare)

    'Return the URL
    BuildGTWebQueryString = sURL
    
End Function

Function fBuildQueryComponent(ByRef sConst As String _
                            , ByRef sField As String) As String
'Called from BuildGTWebQueryString to help build the URL,
' by returning a null string of the field is empty,
' or adding the appropriate prefix when the field contains a value

    If Len(sField) = 0 Then
        fBuildQueryComponent = vbNullString
    Else
        fBuildQueryComponent = sConst & sField
    End If
    
End Function

Sub SetUpQuerySpecSheetWeb(ByRef wksQuerySpec As Worksheet _
                         , ByRef vQueryList() As Variant _
                         , ByRef vRegions() As Variant _
                         , ByRef vURL As Variant _
                         , StartTime As Date)
'Create a worksheet that summarises the query specification and the extraction process
' This allows later auditing of what was done.
    
    Dim rng As Range
    Dim iLastRow As Integer
    Dim vTransposeResult() As Variant           'Created this 2020-09 to use in the TransposeArray output
    
    With wksQuerySpec
        .Name = "Query specification"
        
        sLogEntry = "Google Trends Web Data API extraction"
        
        With .Cells(1, 1)
            .Value = sLogEntry
            .Style = "Heading 1"
        End With
        sLogEntry = sLogEntry & " (" & fCurrentVersionNumber & "),"
        
        'List the query terms
        With .Cells(2, 1)
            .Value = "Query term" & IIf(LCase$(Left(fvReturnNameValue("API_TermForMethod", bCheckForActiveWorkbook:=False), 5)) = "Graph", "s", vbNullString)
            .Style = "Heading 2"
        End With
        
        'Write the query list to the worksheet
        iLastRow = Application.WorksheetFunction.Max(21, UBound(vQueryList) + 3)
        Set rng = .Cells(3, 1)
        '2020-09: Not using TransposeArray for a single dimension array
        rng.Resize(UBound(vQueryList), 1).Value = Application.Transpose(vQueryList)
        sLogEntry = sLogEntry & "Query Terms:" & fReturnQueryListAsOneString(sWorksheetPrefix:="W") & ","
        
        'Copy the date specification
        CopyRange rngTarget:=.Range(.Cells(2, 3), .Cells(4, 4)) _
                , rngSource:=Sheet11.Range("CompleteDateRangeSpecificationWeb")
        .Cells(2, 3).Style = "Heading 2"
        If Len(.Cells(3, 4).Value) = 0 Then .Cells(3, 4).Value = Format(#1/1/2004#, "yyyy/mm/dd")
        If Len(.Cells(4, 4).Value) = 0 Then .Cells(4, 4).Value = Format(Now, "yyyy/mm/dd")
        
        sLogEntry = sLogEntry & "Start Date:" & Format(fvReturnNameValue("WStartDate"), "yyyy-mm-dd") & ","
        sLogEntry = sLogEntry & "End Date:" & Format(fvReturnNameValue("WEndDate"), "yyyy-mm-dd") & ","
        
        'Copy the geographic specification
        CopyRange rngTarget:=.Range(.Cells(7, 3), .Cells(11, 4)) _
                , rngSource:=Sheet11.Range("CompleteLocationRangeSpecificationWeb")
        Select Case fvReturnNameValue("WGeographicLevel")
        Case "Worldwide"
            .Range(.Cells(9, 4), .Cells(10, 4)).ClearContents
        Case "Country"
            .Cells(10, 4).ClearContents
        Case "Region"
            'Possible to-do item (see BuildGTWebQueryString as well):
            'Modify this if a multi-region search is employed for GetGraph
'            'If all regions are to be queried (i.e., Region was left blank), then write all regions to the sheet
'            If UBound(vRegions) > 1 Then
'                Set rng = .Cells(10, 4)
'                rng.Resize(1, UBound(vRegions)).value = Application.Transpose(vRegions)
'            End If
        End Select
        .Cells(7, 3).Style = "Heading 2"

        sLogEntry = sLogEntry & "Geographic Level:" & fvReturnNameValue("WGeographicLevel") & ","
        sLogEntry = sLogEntry & "Country:" & fvReturnNameValue("WCountry") & ","
        sLogEntry = sLogEntry & "Region:" & fvReturnNameValue("WRegion") & ","
        sLogEntry = sLogEntry & "Descriptor:" & fvReturnNameValue("WDescriptor") & ","

        'Copy the Function specification
        CopyRange rngTarget:=.Range(.Cells(13, 3), .Cells(13, 4)) _
                , rngSource:=Sheet11.Range("CompleteFunctionRangeSpecification")
        .Cells(13, 3).Style = "Heading 2"
        sLogEntry = sLogEntry & "Function:" & fvReturnNameValue("WFunction") & ","
        
        'Copy the Domain specification
        CopyRange rngTarget:=.Range(.Cells(15, 3), .Cells(15, 4)) _
                , rngSource:=Sheet11.Range("CompleteDomainRangeSpecification")
        .Cells(15, 3).Style = "Heading 2"
        sLogEntry = sLogEntry & "Search domain:" & fvReturnNameValue("WDomain") & ","
        
        'Copy the Search Category specification
        CopyRange rngTarget:=.Range(.Cells(17, 3), .Cells(17, 4)) _
                , rngSource:=Sheet11.Range("CompleteCategoryRangeSpecification")
        .Cells(17, 3).Style = "Heading 2"
        If Len(.Cells(17, 4).Value) = 0 Then .Cells(17, 4).Value = "All"
        sLogEntry = sLogEntry & "Search category:" & fvReturnNameValue("WCategory") & ","
        
        'Set the date and time of extraction
        With .Cells(2, 7)
            .Value = "Date and time of extraction:"
            .Style = "Heading 2"
        End With
        .Cells(3, 7).Value = "Start:"
        With .Cells(3, 8)
            .Value = StartTime
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
        With .Cells(19, 3)
            .Value = "Sample URL:"
            .Style = "Heading 2"
        End With
        'Do not store the API Key in the logged URL
        .Cells(19, 4).Value = Left(vURL, InStr(1, vURL, "&key=") + 4) & sReqStrGTEH_APIKey
        
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
        sLogEntry = sLogEntry & "Extraction started:" & Format(StartTime, "yyyy-mm-dd hh:mm:ss") & ","
        sLogEntry = sLogEntry & "Extraction completed:" & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        
        WriteLogEntry iLogEvent:=LogEventGTw, sLogEntry:=sLogEntry
        
        'Set the time for the completion of the analysis
        ' This can be done here, because this is the last step in the process
        .Cells(4, 8).Value = Now
    
        'Do a dump of the raw query specification information for robust later retrieval
        WriteQSvaluesToFile wbkTarget:=wksQuerySpec.Parent _
                          , wksSpecToStore:=Sheet11

    End With 'wksQuerySpec
End Sub

Function fDoAllErrorCheckingForWeb() As Boolean
'This function does all the error checking across each of the input areas.
' When any error is found, it alerts the user and ends execution.
' If then end of the function is reached, then theoretically, no errors were encountered.

    'Recalculate the sheet to ensure the on-sheet error checking is up to date
    Sheet11.Calculate

    'Check for any error messages. This is the first level check, which needs to be passed.
    'There is one error message starting with "Note:"
    If fvReturnNameValue("NErrorMessagesWeb") > 1 _
        Or Not (fvReturnNameValue("NErrorMessagesWeb") = 1 _
        And (Left(fvReturnNameValue("WErrorDisplay1"), 4) = "Note" _
        Or Left(fvReturnNameValue("WErrorDisplay1"), 3) = "All")) Then
        MsgBox "There are still outstanding input error messages that need to be resolved before the query can be sent." _
            & vbCrLf & "Please attend to these issues and then try again." _
            , vbCritical + vbOKOnly, "Input Error Messages outstanding"
        EndGracefully
    End If

    'Now, even though it is a repetition of the UI error checking on the worksheet,
    ' as a failsafe, each input field is checked in much the same way as on the UI
    
    'These defined names check for completion of various parts of the UI input
    '=CompletedAllW
    '=CompletedAPIKey
    '=CompletedFunctionW
    '=CompletedDataTargetW
    '=CompletedLocationW
    '=CompletedSearchTermsW
    
    'Check that the function is specified
    If Not fvReturnNameValue("CompletedFunctionW") Then
        Range("WFunction").Select
        MsgBox "You must specify the function to use!" _
            , vbCritical + vbOKOnly _
            , "No function specification"
        EndGracefully
    End If
'    fvReturnNameValue ("API_TermForMethod")
    
    'Check that at least one query term exists
    If Not fvReturnNameValue("CompletedSearchTermsW") Then
        Range("WSearchTerm01").Select
        MsgBox "You must specify at least one Query term!" _
            , vbCritical + vbOKOnly _
            , "No query terms"
        EndGracefully
    End If
    
    'Check that the search location has been specified properly
    If Not fvReturnNameValue("CompletedLocationW") Then
        Range("WGeographicLevel").Select
        MsgBox "The geographic location is not specified correctly!" _
            , vbCritical + vbOKOnly _
            , "Incorrection geographic specification"
        EndGracefully
    End If
    'Now check all the possible combinations of the geographic location specification
    'First check the level
    'Check that the level exists
    If Not Len(fvReturnNameValue("WGeographicLevel")) > 0 Then
        Range("WGeographicLevel").Select
        MsgBox "You must specify the Geographic Level!" _
            , vbCritical + vbOKOnly _
            , "No Geographic Level specification"
        EndGracefully
    End If
    'Check that the level is in the allowable list
    If Not fIsInListNamedRange(fvReturnNameValue("WGeographicLevel"), "GeoLevels", False) Then
        Range("WGeographicLevel").Select
        MsgBox "The value '" & fvReturnNameValue("WGeographicLevel") & "' specified for the Geographic level is not a valid entry!" _
            , vbCritical + vbOKOnly _
            , "Incorrect Geographic level specification"
        EndGracefully
    End If
    
    Select Case fvReturnNameValue("WGeographicLevel")
'    Case "Worldwide"
'        'Do nothing further
'    Case Else
    Case "Country", "Region"
        'If the level is Country, check the country value as well
        'If the level is Region, check the country and region values as well
        'So first we check the country, because that must be checked regardless
        'Check that the Country exists
        If Not Len(fvReturnNameValue("WCountry")) > 0 Then
            Range("WCountry").Select
            MsgBox "You must specify the Country!" _
                , vbCritical + vbOKOnly _
                , "No Country specification"
            EndGracefully
        End If
        'Check that the Country is in the allowable list
        If Not fIsInListNamedRange(fvReturnNameValue("WCountry"), "CountryNames", False) Then
            Range("WCountry").Select
            MsgBox "The value '" & fvReturnNameValue("WCountry") & "' specified for the Country is not a valid entry!" _
                , vbCritical + vbOKOnly _
                , "Incorrect Country specification"
            EndGracefully
        End If
        
        'Then if the level is region, we check that too
        If fvReturnNameValue("WGeographicLevel") = "Region" Then
''            'This check is currently disabled. See mGoogleTrendsInfoExtraction.fDoAllErrorChecking
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
            If Not fIsInListNamedRange(fvReturnNameValue("WRegion"), "WCountrySubdivisionDynamicList", False) And Not fvReturnNameValue("WRegion") = vbNullString Then
                Range("WRegion").Select
                MsgBox "The value '" & fvReturnNameValue("WRegion") & "' specified for the Region is not a valid entry!" _
                    , vbCritical + vbOKOnly _
                    , "Incorrect Region specification"
                EndGracefully
            End If
        End If
    End Select
    
    'Check that the start and end dates are valid dates, and that the start date is after the end date
    'These are the UI error checking formulas:
    '=IF(AND(LEN(WStartDate)>0,WStartDate<DATE(2004,1,1)),"The Start Date is earlier than 1 January 2004. No Google Trends data is available before 2004. Please set the Start Date to 2004/1/1 or later.","")
    '=IF(LEN(WEndDate)>0,IF(WEndDate<DATE(2004,1,1),"The End Date is earlier than 1 January 2004. No Google Trends data is available before 2004. Please set the End Date to 2004-01 or later.",IF(WEndDate<WStartDate,"The End Date is earlier than the Start Date. This is an invalid time specification. Please set the End Date to "&TEXT(WStartDate,"yyyy-mm")&" or later.","")),"")
    
    'Check that the Start Date is a valid date
    If Not IsDate(fvReturnNameValue("WStartDate")) And Not Len(fvReturnNameValue("WStartDate")) = 0 Then
        Range("WStartDate").Select
        MsgBox "The Start Date value of '" & fvReturnNameValue("WStartDate") & "' is not a valid date!" _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the Start Date is not before 2004/1/1
    If fvReturnNameValue("WStartDate") < DateSerial(2004, 1, 1) And Not Len(fvReturnNameValue("WStartDate")) = 0 Then
        Range("WStartDate").Select
        MsgBox "The Start Date is before the 1st of January 2004!" _
            & " There are no Google Trends data before that date." _
            & vbCrLf & "Please set the Start Date to " _
            & Format(IIf(fvReturnNameValue("WStartDate") <= DateSerial(2004, 1, 3), DateSerial(2004, 1, 4), DateSerial(2004, 1, 1)), "yyyy-mm-dd") & "." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the End Date is a valid date
    If Not IsDate(fvReturnNameValue("WEndDate")) And Not Len(fvReturnNameValue("WEndDate")) = 0 Then
        Range("WEndDate").Select
        MsgBox "The End Date value '" & fvReturnNameValue("WEndDate") & " is not a valid date!" _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the End Date is not later than the present date
    If fvReturnNameValue("WEndDate") > Now() Then
        Range("WEndDate").Select
        MsgBox "The End Date " & Format(fvReturnNameValue("WEndDate"), "yyyy-mm-dd") & " is in the future (today's date is " & Format(Now(), "yyyy-mm-dd") & ")!" _
            & " Google Trends web data are only available for historical dates." _
            & vbCrLf & "Please set the End Date to " & Format(Now(), "yyyy-mm-dd") & " or earlier, or leave it blank for up-to-the present data." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
        EndGracefully
    End If
    
    'Check that the Start Date is not after the End date
    If fvReturnNameValue("WStartDate") > fvReturnNameValue("WEndDate") And Not Len(fvReturnNameValue("WEndDate")) = 0 Then
        Range("WStartDate").Select
        MsgBox "The Start Date (" & Format(fvReturnNameValue("WStartDate"), "yyyy-mm-dd") & ") is after the End Date (" & Format(fvReturnNameValue("WEndDate"), "yyyy-mm-dd") & ")!" _
            & vbCrLf & "This is an invalid time specification." _
            & vbCrLf & "Please set the Start Date to before the End Date, or the End Date to after the Start Date." _
            , vbCritical + vbOKOnly _
            , "Incorrect date specification"
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
    If Not fvReturnNameValue("CompletedDataTargetW") Then
        Range("WDataTarget").Select
        MsgBox "No file and path name has been specified for the output file!" _
            & vbCrLf & "Please double click on the cell to specify the file name and location for the output file." _
            , vbCritical + vbOKOnly _
            , "No output file"
        EndGracefully
    End If

    'Check that the specified directory is accessible
    If Not FileDirCheck(2, fvReturnNameValue("WDataTarget")) Then
        Range("WDataTarget").Select
        MsgBox "The directory specified for the data extraction file cannot be found!" _
            & vbCrLf & "Please double click on the cell to specify the file name and location for the output file." _
            , vbCritical + vbOKOnly _
            , "Directory not found"
        EndGracefully
    End If
    
    'Next check that the file does not already exist
    If Len(Dir(fvReturnNameValue("WDataTarget"), vbNormal)) > 0 Then
        Range("WDataTarget").Select
        MsgBox "The file specified for the output file already exists!" _
            & vbCrLf & "Please double click on the cell to specify a new file for the output." _
            , vbCritical + vbOKOnly _
            , "Output file already exists"
        EndGracefully
    End If
    
    'If this point is reached, and no error check has bombed out, then we assume(!) there are no errors
    fDoAllErrorCheckingForWeb = True

End Function

