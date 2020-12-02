Attribute VB_Name = "mRequestAndParseData"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the code which does the actual extraction and parsing      '
' of data from the Google Trends Extended for Health (a.k.a. Google Flu) service. '
' It calls procedures from several of the other modules.                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'These two variables allow the code to attempt to recover from an error that is returned by the API,
' but keep track of whether it is the first or successive times that an error has been raised,
' so that an API error leads to a repeat of the request to the API,
' in the hope that temporary connection errors do not lead to the whole process failing
Dim vAccumulatedError() As Variant
Dim iAccumulatedErrorCount As Integer
Sub workonjson()
fvGetAndParseGoogleData "abc", Date, 1
End Sub
Function fvGetAndParseGoogleData(ByVal sURL As String _
                                        , ByRef sDate As Date _
                                        , ByRef iNTerms As Integer) As Variant
'Sends the API request URL to Google using the fGetData function,
' and then parses the returned JSON string into a variant
'sURL is passed a string from vURLArray(j), the variant built by fBuildRequestArray
'The array returned by fvGetAndParseGoogleData is two-dimensional.
' The structure of the array is counter-intuitive, as only the second dimension can be resized using ReDim Preserve.
' There are as many columns (2nd dimension) as the number of requested dates + 1.
'  The values in the first column are the terms, and the values in subsequent columns are the Google Trends values.
' There are as many rows as the number of terms + 1.
' There is one row for the dates and one subsequent row for each term (iNTerms).

    Dim V() As Variant              'Variant that contains the extracted data from the JSON string
    Dim sResult As String           'Captures the data returned by the Google Trends server in the HTTP request
    Dim sSearchString1 As String    'Element to search for in JSON string
    Dim sSearchString2 As String    'Element to search for in JSON string
''    Dim sTerm As String             'the query term to search for in the JSON string
    Dim lPISstart As Long           'A counter to determine the starting point for the data being extracted from the JSON string
    Dim lPISend As Long             'A counter to determine the end point for the data being extracted from the JSON string
''    Dim sSearchString As String     'What to search for in the JSON string
''    Dim iCounter As Integer         'A counter that loops through elements of the JSON string
'Replaced with i now that parsing is done externally
    Dim i As Long
''    Dim StartDate As Date           'The date retrieved from the JSON string
'I don't like this, but I am using byRef to pass StartDate back to the calling sub
    Dim aTermPoints() As Long       'Stores the points in the JSON string at which the phrase "term: " is found
                                    ' so that various terms in the query can be parsed out
''    Dim iTermCounter As Integer     'Count which term is being extracted
'Replaced with j now that parsing is done externally
    Dim j As Long
    Dim bDateCheck As Boolean       'Test whether the dates created by the program checks out with the dates in the JSON string
    
    Dim dicResult As Object         'Added 2020-08-07 to parse JSON data. To be late-bound below
    Dim bYearlyData As Boolean          'Check whether the JSON string contains year values (where only the year number is supplied) or not (full dates supplied for monthly and daily values)
    Dim vTerms As Variant               'Array to store all terms for which search volumes were retrieved
    Dim vDates As Variant               'Array to store all dates in returned sampling
    Dim vValues As Variant              'Array of all values for any particular term, for all dates in the sampling
    
    'Pass the URL to Google and get the result back:
    sResult = fGetData(sURL)
    'For testing, write the returned string to a text file
    'At the moment, I am not testing whether fWriteToFile returns False as an indication of errors
    '2020-10-28 modified fWriteToFile to onclude bAppend so that all json strings returned can be added incrementally to the file
    If Not bProductionVersion Then _
        fWriteToFile sFile:=Replace(fvReturnNameValue(sName:="DataTarget", bCheckForActiveWorkbook:=False), ".xlsx", ".json", , , vbTextCompare) _
                   , sFileContents:=sResult _
                   , bAppend:=True
    
    'If the quota exceeded error is returned, then there is no data to parse, so exit the function
    ' bQuotaExceeded is declared at the module level and set in fGetData
    If bQuotaExceeded Then
        'Fill fvGetAndParseGoogleData with a dummy value so that an error is not generated on return
        ReDim V(1 To 1)
        V(1) = sColon
        fvGetAndParseGoogleData = V
        Exit Function
    End If
    
    '2020-08-07
    'Parse the JSON string using Daniel Ferry's JSON parser
    Set dicResult = CreateObject("Scripting.Dictionary")
    Set dicResult = ParseJSON(sResult)

    bYearlyData = InStr(1, sURL, "Resolution=Year", vbTextCompare) > 0
    
    'Get the list of terms from the JSON string
    vTerms = GetFilteredValues(dicResult, "*.term*")
    'Get the dates and values for each term from the JSON string
    vDates = GetFilteredValues(dicResult, "*.lines(0).*date*")
    'Dimension v to contain the data
    'First dimension is one more than the number of terms (first column contains dates)
    'Second dimension is one more than the number of values (first row contains 'titles': 'date' and each term
    ReDim V(1 To UBound(vTerms) + 1, 1 To UBound(vDates) + 1)

    'This was done in my own parsing code, moved here 2020-09-03 because parsing is not completing using Daniel Ferry's parser
    'set sDate to the first date value
    'Because sDate is passed ByRef, this stores the date for later use in the calling procedure
    sDate = vDates(LBound(vDates))
    
    'Add the date label and date values to V
    V(1, 1) = "date"
    For i = LBound(vDates) To UBound(vDates)
        'If yearly data are requested, DateValue cannot be used, so add the month/day
        If bYearlyData Then
            V(1, i + 1) = DateSerial(vDates(i), 1, 1)
        Else
            V(1, i + 1) = DateValue(vDates(i))
        End If
        'Error check the date calculation
        'Stricly speaking, all errors in the process should have been intercepted in fGetData (fTestHTTPDataForErrors)
        If Err.Number <> 0 Then
            Err.Clear
    '       Debug.Print Mid$(sResult, lPISstart, lPISend - lPISstart)
            MsgBox "An error occured when attempting to interpret the date " _
                & vDates(i) _
                & " found for value [" & i & "]" _
                & " from the JSON string returned by Google Trends." _
                & vbCrLf & "Please check the query specification and retry the process." _
                , vbCritical + vbOKOnly, "Invalid date found"
            
            'Do not end completely, but exit the function, so that existing data can be reported
            Exit Function
        End If
    Next i

    'Retrieve each set of values and write to their term
    For i = LBound(vTerms) To UBound(vTerms)
        'Add the terms to the first row
        V(i + 1, 1) = vTerms(i) 'i+1 to make place for the date column
        'Read the values and add them to the remainder of the column array dimension
        vValues = GetFilteredValues(dicResult, "*.lines(" & i - 1 & ").*value*")    'i-1 because the dictionary is zero-based
        For j = LBound(vValues) To UBound(vValues)
            V(i + 1, j + 1) = vValues(j)
        Next j
    Next i
'--- This code is commented out, as it is obviated by the use of Daniel Ferry's JSON parser ---'
''''GoTo SkipAllTheOldCode
''''
''''    'Parse the JSON string returned by Google Trends
''''    iTermCounter = 0
''''    lPISstart = 1
''''
''''    'Parse multiple query terms
''''    'To do this, first check the number of iQTCount
''''    'Then build an array storing the start point for each series (the position where the word "term" is found)
''''    ' (this array stores lPISstart and lPISend for each instance of a term)
''''    'Then loop through the dates while lPISstart < the next term's starting position,
''''    ' as stored in the array, or < Len(sResult) for the last one
''''    iNTerms = UBound(vTerms)
''''    ReDim V(1 To iNTerms + 1, 1 To 1)       'One element is added for the dates
''''    V(1, 1) = "date"
''''    ReDim aTermPoints(1 To iNTerms + 1)     'One more is added, so that an end point beyond the last term can be captured
''''
''''    'Loop through the finds in the string
''''    'First find the term
''''    sSearchString1 = sQuote & "term" & sQuote & ": " & sQuote
''''    lPISstart = InStr(lPISstart, sResult, sSearchString1, vbTextCompare) + Len(sSearchString1)
''''    'If an error string is returned, there will be no instance of {"term": "___" in the string
''''    If lPISstart = 0 Then
'''''        Debug.Print Left$(sResult, 1000)
''''        MsgBox "The JSON string returned by Google Trends does not contain data points for a search term." _
''''            & vbCrLf & "The string (or part thereof) will now be displayed." _
''''            , vbCritical + vbOKOnly, "No term found"
''''        MsgBox sResult, vbInformation + vbOKOnly, "Google Trends JSON result"
''''        Exit Function
''''    Else    'No error, so store the start points for each term in aTermPoints
''''        lPISend = InStr(lPISstart, sResult, sQuote, vbTextCompare)
''''        V(2, 1) = Mid(sResult, lPISstart, lPISend - lPISstart)
''''        'Store the position where this query term was found for later extraction
''''        aTermPoints(1) = lPISstart
''''        iTermCounter = 1
''''
''''        'Now loop through the JSON string and extract the remaining terms
''''        Do While lPISstart <> 0 'And lPISstart < Len(sResult)
''''            lPISstart = InStr(lPISstart, sResult, sSearchString1, vbTextCompare)
''''            If lPISstart > 0 Then
''''                lPISstart = lPISstart + Len(sSearchString1)
''''                iTermCounter = iTermCounter + 1
''''            Else
''''                Exit Do
''''            End If
''''            lPISend = InStr(lPISstart, sResult, sQuote, vbTextCompare)
''''            V(iTermCounter + 1, 1) = Mid(sResult, lPISstart, lPISend - lPISstart)
''''            'Store the position where this query term was found for later extraction
''''            aTermPoints(iTermCounter) = lPISstart
''''        Loop
''''
''''        If iTermCounter <> iNTerms Then
''''            'Not all the terms could be found
''''            MsgBox "There was an error returning all the data from the Google Trends query. " _
''''                & "Not all the terms were returned.", vbCritical + vbOKOnly, "Incorrect data string returned"
''''
''''            EndGracefully
''''        End If
''''    End If
''''    'To find the values for the last term, an additional point for the last term must be set to the very end of the result string
''''    aTermPoints(UBound(aTermPoints)) = Len(sResult)
''''
''''    'Reset the counters
''''    lPISstart = 1
''''
''''    'Set the strings in the JSON code which serve as target points
''''    sSearchString1 = sQuote & "date" & sQuote & ": " & sQuote
''''    sSearchString2 = sQuote & "value" & sQuote & ": "
''''
''''    'Extract the data for each search term
''''    For iTermCounter = 1 To iNTerms
''''        'Reset counter
''''        iCounter = 0    'Set iCounter to 0 for each iteration
''''        'Set lPIStart to the point where the term for this iteration was found
''''        ' (it will have exceeded that in the preceding iteration, so needs to be brought back)
''''        lPISstart = aTermPoints(iTermCounter)
''''        'Loop through the JSON string and extract all the values and the dates
''''        Do While lPISstart <> 0 'And lPISstart < Len(sResult)
''''            'Find the date
''''            lPISstart = InStr(lPISstart, sResult, sSearchString1, vbTextCompare)
''''            If lPISstart > 0 And lPISstart < aTermPoints(iTermCounter + 1) Then
''''                lPISstart = lPISstart + Len(sSearchString1)
''''                iCounter = iCounter + 1
''''                'On the first iteration, incrementally increase the size of the array
''''                If iTermCounter = 1 Then _
''''                    ReDim Preserve V(1 To iNTerms + 1, 1 To iCounter + 1)
''''            Else
''''                Exit Do
''''            End If
''''
''''            lPISend = InStr(lPISstart, sResult, sQuote, vbTextCompare)
''''            On Error Resume Next
''''                If iTermCounter = 1 Then
''''                'For the first round, store the dates in the first element of the array
''''                    'If yearly data are requested, DateValue cannot be used, so add the month/day
''''                    If InStr(1, sURL, "Resolution=Year", vbTextCompare) > 0 Then
''''                        V(1, iCounter + 1) = DateSerial(Mid(sResult, lPISstart, lPISend - lPISstart), 1, 1)
''''                    Else
''''                        V(1, iCounter + 1) = DateValue(Mid(sResult, lPISstart, lPISend - lPISstart))
''''                    End If
''''                    'Error check the date calculation
''''                    'Stricly speaking, all errors in the process should have been intercepted in fGetData (fTestHTTPDataForErrors)
''''                    If Err.Number <> 0 Then
''''                        Err.Clear
'''''                        Debug.Print Mid$(sResult, lPISstart, lPISend - lPISstart)
''''                        MsgBox "An error occured when attempting to interpret the date " _
''''                            & Mid$(sResult, lPISstart, lPISend - lPISstart) _
''''                            & " found for value [" & iTermCounter & "," & iCounter & "]" _
''''                            & " from the JSON string returned by Google Trends." _
''''                            & vbCrLf & "Please check the query specification and retry the process." _
''''                            , vbCritical + vbOKOnly, "Invalid date found"
''''
''''                        'Do not end completely, but exit the function, so that existing data can be reported
''''                        Exit Function
''''                    End If
''''                Else
''''                'For subsequent iterations (i.e., term2...termN), just check the date against what is stored
''''                'In theory, this should never be triggered, because it means there was some glitch in the data returned by the Google API
''''                    bDateCheck = True
''''                    If InStr(1, sURL, "Resolution=Year", vbTextCompare) > 0 Then
''''                        If V(1, iCounter + 1) <> DateSerial(Mid(sResult, lPISstart, lPISend - lPISstart), 1, 1) Then bDateCheck = False
''''                    Else
''''                        If V(1, iCounter + 1) <> DateValue(Mid(sResult, lPISstart, lPISend - lPISstart)) Then bDateCheck = False
''''                    End If
''''
''''                    If Not bDateCheck Then _
''''                    MsgBox "The date (" & Mid(sResult, lPISstart, lPISend - lPISstart) & ") in row " _
''''                        & iCounter + 1 & " for term " & iTermCounter + 1 & "(" & V(iTermCounter + 1, 1) _
''''                        & ") does not correspond to the same date (" & V(1, iCounter + 1) _
''''                        & ") for term 1 (" & V(2, iCounter + 1), vbInformation + vbOKOnly _
''''                        , "Date discrepancy"
''''
''''                End If
''''            On Error GoTo 0
''''
''''            'For the first round, find and test the first date value, and set it to sDate
''''            'Because sDate is passed ByRef, this stores the date for later use in the calling procedure
''''            If iTermCounter = 1 And iCounter = 1 Then sDate = V(iTermCounter, iCounter + 1)
''''
''''            'Get the value
''''            lPISstart = InStr(lPISstart, sResult, sSearchString2, vbTextCompare) + Len(sSearchString2)
''''            lPISend = InStr(lPISstart, sResult, "}", vbTextCompare)
''''            V(iTermCounter + 1, iCounter + 1) = CDbl(Mid(sResult, lPISstart, lPISend - lPISstart))
''''        Loop
''''    Next iTermCounter    'iTermCounter = 1 to iNTerms
''''SkipAllTheOldCode:
    'When all data have been parsed from the string, pass the temp array to fvGetAndParseGoogleData for processing in the calling procedure
    ' v is a two-dimensional array and contains the search term as its first element
    fvGetAndParseGoogleData = V
    
End Function

Function fGetData(ByRef sURL As String) As String
'Modified from code found at:
'http://tkang.blogspot.com.au/2010/09/sending-http-post-request-with-vba.html
'Pass a HTTP request to Google, and test the result for some API error values by calling fTestHTTPDataForErrors

'First clear the error recording variables
iAccumulatedErrorCount = 0              'Declared at module level
ReDim vAccumulatedError(1 To 2, 1 To 1) 'Using ReDim without Preserve clears the array

Dim iReturnError As Integer             'Counts the number of errors returned in successive HTTP requests

Dim oWinHttpRequest As Object
Set oWinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
     
    Do
        oWinHttpRequest.Open "GET", sURL, False
        On Error Resume Next
        oWinHttpRequest.Send
''        If Err.Number <> 0 Then
''            If Err.Number = -2147012889 Then
''                'The server name or address could not be resolved (80072ee7)
''                MsgBox "There appears to be an internet connection error." _
''                    & vbCrLf & "Please check your internet connection and attempt the data extraction again." _
''                    , vbCritical + vbOKOnly, "Server/address not resolved"
''
''                EndGracefully
''            Else
''                Err.Raise Err.Number
''            End If
''            Err.Clear
''
''        End If
''
''        'Store the response from the HTTP request
''        fGetData = oWinHttpRequest.responseText
        
        If Err.Number = 0 Then
            'No error was raised in sending the request, so get the response text
            'If an error was raised, it will be dealt with below
            'Store the response from the HTTP request
            fGetData = oWinHttpRequest.responseText
        End If

        'Added 2020-10-01
        'No data is returned.
        'to-do: consider still creating the output workbook showing no data
        If fGetData = "{}" Then
            MsgBox "The Google Trends service returned an empty JSON string!" _
                , vbCritical + vbOKOnly, "No data returned"
            End
        End If
        
        'If an error is returned, alert the user, and count the error, so that the HTTP request can be resent in the hope of getting a result
        If Err.Number <> 0 Then
            'First deal with the error raised during sending
            If Err.Number = -2147012889 Then
                'The server name or address could not be resolved (80072ee7)
                MsgBox "There appears to be an internet connection error." _
                    & vbCrLf & "Please check your internet connection and attempt the data extraction again." _
                    , vbCritical + vbOKOnly, "Server/address not resolved"
                EndGracefully
            Else
                'Next look at possible errors raised in receiving the response
                iReturnError = iReturnError + 1
                If iReturnError = 1 Then
                    MsgBox "The following error was raised:" _
                    & vbCrLf & Err.Number _
                    & vbCrLf & Err.Description _
                    & vbCrLf & "Please click OK to attempt to extract the sampling again." _
                    , vbInformation + vbOKOnly, "Error in html response"
                    Err.Clear
                Else    'This is not the first error, but still give the option of sending the request again
                    If MsgBox("The following error was raised:" _
                    & vbCrLf & Err.Number _
                    & vbCrLf & Err.Description _
                    & vbCrLf & "Do you want to attempt to extract the sampling again (OK), or cease the extraction process completely (Cancel)?" _
                    , vbInformation + vbOKCancel, "Continued error in html response") = vbCancel Then
                        
                        ''EndGracefully
                        'Rather than ending, exit the function and pass the error back as well (via bOtherHTTPError) so that the results which have already been obtained can be cleaned up)
                        Err.Clear
                        bOtherHTTPError = True
                        fGetData = vbNullString
                        Exit Function
                    Else
                        Err.Clear
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    'The loop will allow me to try again if certain errors are encountered
    Loop While fTestHTTPDataForErrors(sDataString:=fGetData) = True
    
End Function

Function fTestHTTPDataForErrors(ByRef sDataString As String) As Boolean
'Tests the returned JSON string for possible API error values.
'Will allow up to three repeats of the HTTP request to see if the error can be resolved.
'This will happen if different errors are returned, or if the same error is returned twice, processing will end.
'Some errors (e.g., invalid API key, 'Not Found', quota exceeded) will stop processing without repeats.
    
    'The Not Found error will not change with repeated requests, so end processing
    If Len(sDataString) >= 9 And Left(sDataString, 9) = "Not Found" Then
        MsgBox "Google returned a 'Not Found' response to this query. No data can be extracted." _
            , vbExclamation + vbOKOnly, "Not Found"
        
            ''EndGracefully
            'Rather than ending, exit the function and pass the error back as well (via bOtherHTTPError) so that the results which have already been obtained can be cleaned up)
            bOtherHTTPError = True
            fTestHTTPDataForErrors = False  'Set this to false, so that continued looping for testing is disabled, but bOtherHTTPError will cause the proces to stop
            Exit Function

    ElseIf InStr(1, sDataString, "keyInvalid", vbTextCompare) > 0 Then
    'The API key is incorrect
        MsgBox "Google has indicated that you are using an invalid API key." _
            & vbCrLf & "Please check the API key in the file '" & fvReturnNameValue("APIKey") & "'." _
            , vbExclamation + vbOKOnly, "Not Found"
        
            ''EndGracefully
            'Rather than ending, exit the function and pass the error back as well (via bOtherHTTPError) so that the results which have already been obtained can be cleaned up)
            bOtherHTTPError = True
            fTestHTTPDataForErrors = False  'Set this to false, so that continued looping for testing is disabled, but bOtherHTTPError will cause the proces to stop
            Exit Function
    
    ElseIf InStr(1, sDataString, "error", vbTextCompare) > 0 Then
    'An error was returned, so parse it and handle it
        
        'First set the error recording to true, so that the do loop in fGetData is repeated
        fTestHTTPDataForErrors = True
        DoEvents   'Return control to Excel
        
        'The first error to test for is if the daily limit has been exceeded
        If InStr(1, sDataString, "Daily Limit Exceeded", vbTextCompare) > 0 Then
            Dim dTime As Date
            Dim iURLStart As Integer, iURLEnd As Integer
            Dim sHelpURL As String
            'Capture the time at which the error was raised, so that if the message box sits there for a long time
            ' (e.g., the computer is unattended), the delay is not unnecessarily extended
            dTime = Now
            iURLStart = InStr(1, sDataString, "https://", vbTextCompare)
            iURLEnd = InStr(iURLStart, sDataString, sQuote, vbTextCompare)
            sHelpURL = Mid(sDataString, iURLStart, iURLEnd - iURLStart)
            
            'The quota is reset at 5 PM Pacific time, I think, so a 24 H wait may not be necessary
''            If MsgBox("You have exceeded your daily Google Trends quota." _
''                & vbCrLf & "No further data can be downloaded today." _
''                & vbCrLf & "Please consult your API console at: " & sHelpURL _
''                & vbCrLf & vbCrLf & "Do you want the program to wait for 24 hours and then try again " _
''                & "(e.g., if this is a dedicated data processing pc)?" _
''                & vbCrLf & "You may not be able to use Excel in that period, depending on the version you have." _
''                , vbCritical + vbYesNo, "Google Trends quota exceeded") = vbYes Then
''                'Wait one day and one second
''                Application.Wait (dTime + 1 + TimeValue("00:00:01"))
''            Else
''                'Turn off the looping of error testing, so that control is returned to the main program
''                bQuotaExceeded = True
''                fTestHTTPDataForErrors = False
''                Exit Function
''            End If
            MsgBox "You have exceeded your daily Google Trends quota." _
                & vbCrLf & "No further data can be downloaded until the quota is reset (at 5PM Pacific time)." _
                & vbCrLf & "Please consult your API console at: " & sHelpURL _
                , vbCritical + vbOKOnly, "Google Trends quota exceeded"
            'Turn off the looping of error testing, so that control is returned to the main program
            bQuotaExceeded = True
            fTestHTTPDataForErrors = False
            Exit Function
            
        ElseIf fExtractPartFromString(sDataString, "code") = "400" Then
            bInvalidArgumentError = True
            fTestHTTPDataForErrors = False
            MsgBox fReturnErrorString(sDataString, True), vbCritical + vbOKOnly, "Invalid query string for Google Trends API"
            Exit Function
            
        ElseIf iAccumulatedErrorCount = 0 Then
        'If the error was not exceeding the quota, record the error,
        ' and then let the do loop from fGetData try again, after a short break
            iAccumulatedErrorCount = 1
            ReDim vAccumulatedError(1 To 2, 1 To 1)
            vAccumulatedError(1, 1) = fExtractPartFromString(sDataString, "code")
            vAccumulatedError(2, 1) = 1
            Application.Wait (Now + TimeValue("0:00:05"))
        
        ElseIf iAccumulatedErrorCount = 1 Then
        'One error has already been returned.
        'Test to see if the second error is the same as the first
            If fExtractPartFromString(sDataString, "code") = vAccumulatedError(1, 1) Then
            'This is the second time this error is being returned
                vAccumulatedError(2, 1) = 2
                'Return an error message and end
                MsgBox fReturnErrorString(sDataString, True), vbCritical + vbOKOnly, "Error obtaining data from Google Trends"
                
                ''EndGracefully
                'Rather than ending, exit the function and pass the error back as well (via bOtherHTTPError) so that the results which have already been obtained can be cleaned up)
                bOtherHTTPError = True
                fTestHTTPDataForErrors = False  'Set this to false, so that continued looping for testing is disabled, but bOtherHTTPError will cause the proces to stop
                Exit Function
            Else
            'This is the second error, but it is different to the first
            'Log the error, and try once more...
                iAccumulatedErrorCount = 2
                ReDim Preserve vAccumulatedError(1 To 2, 1 To 2)
                vAccumulatedError(1, 2) = fExtractPartFromString(sDataString, "code")
                vAccumulatedError(2, 2) = 1
            End If
        ElseIf iAccumulatedErrorCount = 2 Then
        'Three errors have been returned.
            'Test to see whether this error is the same as one of the previous errors
            If fExtractPartFromString(sDataString, "code") = vAccumulatedError(1, 1) Then
                vAccumulatedError(2, 1) = vAccumulatedError(2, 1) + 1
            ElseIf fExtractPartFromString(sDataString, "code") = vAccumulatedError(1, 2) Then
                vAccumulatedError(2, 2) = vAccumulatedError(2, 2) + 1
            Else
                ReDim Preserve vAccumulatedError(1 To 2, 1 To 3)
                vAccumulatedError(1, 3) = fExtractPartFromString(sDataString, "code")
                vAccumulatedError(2, 3) = 1
            End If
            
            Dim sErrStringThrice As String
            Dim d As Integer
            
            sErrStringThrice = "Three errors have occurred while trying to obtain data from Google Trends." _
                & vbCrLf & "The errors (and number of times encountered) were:"
            
            For d = LBound(vAccumulatedError, 2) To UBound(vAccumulatedError, 2)
                sErrStringThrice = vbCrLf & sErrStringThrice & "Code: " & vAccumulatedError(1, d)
                sErrStringThrice = sErrStringThrice & "(" & vAccumulatedError(2, d) & ")"
            Next d
            
            sErrStringThrice = vbCrLf & "This program will now end."
            MsgBox sErrStringThrice, vbCritical + vbOKOnly, "Repeated errors"
            
            ''EndGracefully
            'Rather than ending, exit the function and pass the error back as well (via bOtherHTTPError) so that the results which have already been obtained can be cleaned up)
            bOtherHTTPError = True
            fTestHTTPDataForErrors = False  'Set this to false, so that continued looping for testing is disabled, but bOtherHTTPError will cause the proces to stop
            Exit Function
        End If
    End If

End Function

Function fReturnErrorString(ByRef sDataString As String, ByVal bWillEnd As Boolean) As String
'Used by fTestHTTPDataForErrors to extract the error from the string returned by the HTTP request
    Dim sErr As String
    sErr = vbNullString
    Call BuildOneErrorStringComponent(sErr, sDataString, "Error Code: ", "code", True)
    Call BuildOneErrorStringComponent(sErr, sDataString, "Domain: ", "domain", True)
    Call BuildOneErrorStringComponent(sErr, sDataString, "Reason: ", "reason", True)
    Call BuildOneErrorStringComponent(sErr, sDataString, "Message: ", "message", True)
    Call BuildOneErrorStringComponent(sErr, sDataString, "Help: ", "extendedHelp", True)
    
    If bWillEnd Then sErr = sErr & "This program will now end."
    fReturnErrorString = sErr
    sErr = vbNullString
End Function

Sub BuildOneErrorStringComponent(ByRef sExistingErrString As String _
                                     , ByRef sDataString As String _
                                     , ByVal sPrefix As String _
                                     , ByVal sSearch As String _
                                     , ByVal isFirstArg As Boolean)
'Used by fReturnErrorString to return one part of the error description
    Dim sComponent As String
    If InStr(1, sDataString, sSearch, vbTextCompare) > 0 Then
        If Not isFirstArg Then sComponent = sExistingErrString & vbCrLf
        sComponent = sComponent & sPrefix '& ": "
        sComponent = sComponent & fExtractPartFromString(sDataString, sSearch)
    Else
        sComponent = vbNullString
    End If
    sExistingErrString = sExistingErrString & sComponent & vbCrLf
End Sub

Function fExtractPartFromString(ByRef sFullString As String, ByRef sNode As String) As String
'Used by BuildOneErrorStringComponent to extract a single portion from the string
    Dim lStart As Long
    Dim lEnd As Long
    Dim sSearchString  As String
    
    sSearchString = sQuote & sNode & sQuote & ": " & IIf(sNode = "code", vbNullString, sQuote)
        
    lStart = InStr(1, sFullString, sNode, vbTextCompare) + Len(sSearchString) - 1
    lEnd = InStr(lStart, sFullString, IIf(sNode = "code", ",", sQuote), vbTextCompare)
    fExtractPartFromString = Mid(sFullString, lStart, lEnd - lStart)
    
End Function
