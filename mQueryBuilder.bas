Attribute VB_Name = "mQueryBuilder"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the functions that build the URLs which will be sent to Google to obtain '
' multiple samples of data (for multiple regions or terms, if requested).                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub test_fBuildRequestArray()
'    fBuildRequestArray
'End Sub

'Function fBuildRequestArray(ByVal sStartTime As String, ByVal sEndTime As String) As Variant
Function fBuildRequestArray() As Variant
'This function returns a list of complete URLs which can be submitted to the API for Google Trends Extended requests
' (i.e., it is not used for GTWeb).
'The number of URLs in the array = the number of date samplings required (=NSamples*2-1) multiplied by the number of regions requested
'For Geographic level as Worldwide or Country, the number of regions is always 1 (at this stage, multi-country comparisons must be done manually)
'If a multi-region request is submitted (i.e., Geographic level is set to "Region" and Region is blank),
' then the list of URL is one for each date sampling, repeated for each region (i.e., region 1, all date samplings, then region 2, all date samplings, etc.)

    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name And Not ActiveSheet.Name = Sheet3.Name Then Exit Function
    
    'General purpose counter variables for this procedure, typically used for setting/looping arrays
    Dim iRegions As Integer
    Dim iDateCounter As Integer
    Dim iRegionCounter As Integer
    Dim iURLCounter As Integer
    
    Dim sTerms As String            'Stores the complete list of search terms specified in the named range SearchTermList
    Dim sTimeRes As String          'Stores the time resolution set in the named range DateResolution
    Dim sKey As String              'Stores the API key when retrieved from the file
    Dim vGeoRes() As Variant        'Declared as variant so that it can store a range of region locations if need be. _
                                      It is always redimmed as two-dimensional, to account for the instance where _
                                      it is fed to a range, which creates a two-dimensional variant.
    Dim vDateArray() As Variant     'The array that stores the start and end dates for each sampling--its upper bound will be the number of samplings
    Dim vRequestArray() As Variant  'The final array which contains a complete URL for each sampling

    sTerms = BuildQueryList(Sheet3)
    'Debug.Print BuildQueryList
    
    sTimeRes = fvReturnNameValue("DateResolution")
    
    Select Case fvReturnNameValue("GeographicLevel")
    Case "Worldwide"
        iRegions = 1
        ReDim vGeoRes(1 To 1, 1 To 1)
        vGeoRes(1, 1) = vbNullString
    Case "Country"
        'Note: At the moment, I am not allowing comparisons across countries,
        ' because the number of samplings * the number of countries (249) can easily exceed the quota
        ' Thus multi-country comparisons must be done manually, with separate queries for each country
        iRegions = 1
        ReDim vGeoRes(1 To 1, 1 To 1)
        vGeoRes(1, 1) = "country=" & fvReturnNameValue("Descriptor")
    Case "Region"
        If fvReturnNameValue("Region") = vbNullString Then
        'No specific region is specified, so search all regions
            'Count the regions
            iRegions = Application.WorksheetFunction.CountIf(Range("ISO3166ParentSubdivisions"), fvReturnNameValue("Descriptor"))
            vGeoRes = Range("GeoAllRegionsForCountry")
            For iRegionCounter = 1 To iRegions
                vGeoRes(iRegionCounter, 1) = "region=" & vGeoRes(iRegionCounter, 1)
            Next iRegionCounter
        Else
        'One specific region is specified
            iRegions = 1
            ReDim vGeoRes(1 To 1, 1 To 1)
            vGeoRes(1, 1) = "region=" & fvReturnNameValue("Descriptor")
        End If
    End Select
    
    'Get the API key value
    sKey = fGetValueFromFile(fvReturnNameValue("APIKey"))

    vDateArray = MergeSamplesAndPeriods
    'The number of requests must be the number of samplings needed * the number of regions requested
    '''ReDim vRequestArray(LBound(vDateArray, 2) To UBound(vDateArray, 2) * UBound(vGeoRes, 1))
    ReDim vRequestArray(LBound(vDateArray, 2) To UBound(vDateArray, 2) * iRegions)
    If UBound(vRequestArray) > iMaxQueriesPD Then
        MsgBox "The number of samplings (" & UBound(vDateArray, 2) _
            & ") multiplied by the number of regions (" & iRegions _
            & ") exceeds the maximum allowable daily quota of " & iMaxQueriesPD & "." _
            & vbCrLf & "Please request the regions one at a time until your quota is exhausted, " _
            & "or reduce the number of samples needed (which might not be advisable if you need accurate results." _
            , vbCritical + vbOKOnly, "Quota will be exceeded"
        EndGracefully
    End If
    'If Regional-comparative requests are made (i.e., a  specific Region is not chosen, then create a url for each region
    ' (All terms for a multi-term query have already been combined in sTerms)
    iURLCounter = 0
    For iRegionCounter = LBound(vGeoRes, 1) To UBound(vGeoRes, 1)
        For iDateCounter = LBound(vDateArray, 2) To UBound(vDateArray, 2)
            iURLCounter = iURLCounter + 1
            
            vRequestArray(iURLCounter) = sReqStrGTEH_URL_Base _
            & sTerms _
            & sReqStrGTEH_StartDate & vDateArray(2, iDateCounter) _
            & sReqStrGTEH_EndDate & vDateArray(3, iDateCounter) _
            & sReqStrGTEH_TimelReso & sTimeRes _
            & IIf(Len(vGeoRes(iRegionCounter, 1)) > 0, sReqStrGTEH_GeoRestr & vGeoRes(iRegionCounter, 1), vbNullString) _
            & sReqStrKey & sKey _
            & sReqStrGTEH_JSON_Request
        
            '2020-10-28 Added a file containing all the URLs when in testing mode
            'N.B. These URLs contain the account key!
            If Not bProductionVersion Then _
                fWriteToFile sFile:=Replace(fvReturnNameValue(sName:="DataTarget", bCheckForActiveWorkbook:=False), ".xlsx", ".URLs", , , vbTextCompare) _
                           , sFileContents:=CStr(vRequestArray(iURLCounter)) _
                           , bAppend:=True
        Next iDateCounter
    Next iRegionCounter
    
    fBuildRequestArray = vRequestArray
    
End Function
Function MergeSamplesAndPeriods() As Variant
'This procedure takes these two function calls:
' mDates.ReturnPeriodsAsArray (returns an array specifying the start and end date for each period)
' mGoogleTrendsInfoExtraction.DrawSamples (returns an array containing an index, and the periods by which each sample is defined)
'and merges them into one so that successive GTeH calls can be made
    Dim arrSamples() As Variant
    Dim arrPeriods() As Variant
    Dim ArrOfSampleDates() As Variant
    Dim i As Integer
    
    'Possible to-do item:
    '   Is the start date or the end date more than Nper from 2004/1/1 or today-2, respectively
    '   [or NPer < the NTotal (i.e., EndDate-StartDate) - NPer]
    '   This will allow the simplified sampling pattern of exending the date ranges beyond the period requested
    '   I can then draw even fewer samples, but I will have to follow this through all the way to where it is written to the worksheet and plotted in the chart
    
    'In the first step, build arrays with an index for each period in each sampling
    arrSamples = mSampling.DrawSamples
    'In the second step, create an array with a start and end date for every period
    arrPeriods = mDates.ReturnPeriodsAsArray(bEndOnError:=True)
    
    'Now each period index is matched up with its corresponding Start- and End dates
    For i = LBound(arrSamples, 2) To UBound(arrSamples, 2)
        arrSamples(2, i) = arrPeriods(1, arrSamples(2, i))
        arrSamples(3, i) = arrPeriods(2, arrSamples(3, i))
    Next i
    
    MergeSamplesAndPeriods = arrSamples

End Function

'Now I can take the merged samples, and use XLWings to run a GT request against each sample, using the specifications from the 'Query speficiation' sheet
'The results of each request must be written to one (numbered--e.g., 'Sampling x') column in a worksheet, and then a last worksheet must be added which averages the values across all the sample worksheets
'Finally, a chart must be plotted from the average worksheet
'Can the average worksheet also calculate a CI for the mean of those values? Will it have to be a surveymean?

Function BuildQueryList(ByRef sht As Worksheet) As String
'Takes each of the specified terms (1..30 terms) and builds a string that can be used in the URL
' The list of terms is added first, so the very first term succeeds the base URL
' (i.e., "https://www.googleapis.com/trends/v1beta/timelinesForHealth?terms=")
' so it does not take the preceding ampersand
'This procedure is called by both the GTe and the GTWeb processes, as the specification for the terms is the same
' (although GTe allows 30 terms, and GTWeb only 5)

    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name And Not ActiveSheet.Name = sht.Name Then Exit Function
    
    Dim i As Integer                'General purpose counter
    Dim j As Integer                'General purpose counter
    Dim sTmp As String              'temp string
    Dim iMaxTerms As Integer        '30 terms for GTe, 5 for GTWeb
    Dim sSheetPrefix As String      '"W" for GTWeb, vbNullString for GTe
    Dim sTermPrefix As String       'The constant sReqStrTerms="&terms=", which is fine for GTe,
                                    ' but must be "&term=" for some applications of GTWeb (see below)
    
    If sht.Name = Sheet3.Name Then
        iMaxTerms = 30
        sSheetPrefix = vbNullString
        sTermPrefix = sReqStrTerms
    ElseIf sht.Name = Sheet11.Name Then
        iMaxTerms = 5
        sSheetPrefix = "W"
        'GTe uses '&terms=', while GTWeb uses '&terms=' for Graph and GraphAverages, and '&term=' for the rest
        If Left(fvReturnNameValue("API_TermForMethod"), 5) <> "graph" Then
            sTermPrefix = Replace(sReqStrTerms, "s", vbNullString, , , vbTextCompare)
        Else
            sTermPrefix = sReqStrTerms
        End If
    End If
    For i = 1 To iMaxTerms
        'First read the cell into sTmp so that it does not have to be read again if it must be processed.
        ' If the cell is empty, sTmp will be a null string
        sTmp = Trim(fvReturnNameValue(sSheetPrefix & "SearchTerm" & Right("0" & CStr(i), 2)))
        If Len(sTmp) > 0 Then
            'Encode characters that are not allowed in URLs
            j = 2
'This was the old method I used, but then I switched to using the EncodeURL worksheet function
'            With Sheet19
'                Do While Len(.Cells(j, 1).value) > 0
'                    If UCase(.Cells(j, 4).value) <> "N" Then _
'                        sTmp = Replace(sTmp, .Cells(j, 2).value, .Cells(j, 3).value, 1, -1, vbTextCompare)
'                    'Debug.Print .Cells(j, 3).value
'                    j = j + 1
'                Loop
'            End With
            
            sTmp = Application.WorksheetFunction.EncodeURL(sTmp)
            'Replace all spaces with '+'
            ' (spaces have already been encoded above)
            sTmp = Replace(sTmp, "%20", "+", 1, -1, vbTextCompare)
            
            'Add the prefix if it is the very first query term
            sTmp = IIf(i = 1, Right(sTermPrefix, Len(sTermPrefix) - 1), sTermPrefix) & sTmp
            'Add the query term to the list of query terms
            BuildQueryList = BuildQueryList & sTmp
        End If
    Next i

End Function

