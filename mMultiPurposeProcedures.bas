Attribute VB_Name = "mMultiPurposeProcedures"
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains a handful of general-purpose subs or functions      '
'  that are called by a wide variety of procedures from all other modules. '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function fvReturnNameValue(ByVal sName As String _
                , Optional bCheckForActiveWorkbook As Boolean = True) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This function is used as a more robust method of returning '
' any value from a named range on any worksheet,            '
' even when that worksheet is not the active sheet,         '
' or from names that are only calculated values             '
'  (i.e., not referring to an actual worksheet range.       '
' Because the interface relies so heavily on names,         '
'  this is a workhorse function which is called many times. '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Under certain conditions This function should fire only when used in this workbook
    If bCheckForActiveWorkbook Then
        If Not ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Function
    Else
        ThisWorkbook.Activate
    End If
    
    On Error Resume Next
    
    fvReturnNameValue = Application.Evaluate(ThisWorkbook.Names(sName).Value)
    
    If Err.Number <> 0 Then
        MsgBox "An error occurred while retrieving the value for the workbook name: '" _
            & sName & "'.", vbCritical + vbOKOnly, "Name error"
        Err.Clear
        EndGracefully
    End If

End Function

Sub StartSmoothly()
''''''''''''''''''''''''''''''''''''''''''''''
' Set some things to enable smoother running '
''''''''''''''''''''''''''''''''''''''''''''''
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub EndGracefully(Optional ByVal e As ErrObject)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Reverse the settings of StartSmoothly,                      '
' and if an error occurred, bring it to the user's attention. '
' Then end code execution.                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Not e Is Nothing Then
'    If e.Number <> 0 Then
        MsgBox "The following error occurred:" _
            & vbCrLf & "Number: " & e.Number _
            & vbCrLf & "Description: " & e.Description _
            , vbCritical + vbOKOnly, "unanticipated Error"
        e.Clear
    End If
    
    With Application
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = vbNullString
    End With
    
    End

End Sub

'Private Sub TestfReturnSafeWorksheetName()
'    fReturnSafeWorksheetName "Sheet1"
'End Sub
Function fReturnSafeWorksheetName(ByRef sStartName As String _
                       , Optional ByRef sExistingName As String _
                       , Optional ByRef wb As Workbook _
                       , Optional ByRef bContinueOnFailure As Boolean) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Return a worksheet name that can be used to name a worksheet without generating an error     '
' bContinueOnFailure is only invoked if no sExistingName is provided.                          '
' If sExistingName is provided and a name cannot be generated, then sExistingName is returned. '
' Makes use of the following two functions below:                                              '
' - fWorksheetNameIsInUse                                                                      '
' - fReturnDefaultOpenWorksheetName                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim i As Integer, j As Integer
    Dim iNCharsToUseInNewName As Integer
    Dim iLoopBailer As Integer
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    If Len(sStartName) = 0 Then
    'No name is specified, so return the default sheet name (in English)
        If Len(sExistingName) = 0 Then
        'If no existing name is provided, then find the next open worksheet in the default sequence of Sheet1, Sheet2, etc.
            fReturnSafeWorksheetName = fReturnDefaultOpenWorksheetName(wb)
        Else
        'If an existing name is provided, then use that (i.e., the worksheet is not renamed)
            fReturnSafeWorksheetName = sExistingName
        End If
    Else
    'A base name is provided
        'First check that it is valid
        'Remove invalid characters
        sStartName = Replace(sStartName, "/", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, "\", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, "[", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, "]", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, "*", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, "?", "_", , , vbTextCompare)
        sStartName = Replace(sStartName, ":", "_", , , vbTextCompare)
        
        'Trim the length to 31 characters
        If Len(sStartName) > 31 Then sStartName = Left(sStartName, 31)
        
        'Check for Reserved names
        If LCase(sStartName) = "history" Then sStartName = "_" & sStartName
        
        'If the name already exists, that means: a) it is valid as is, b) it cannot be used unmodified
        If fWorksheetNameIsInUse(sWorksheetName:=sStartName, wb:=wb) Then
            'Attempt to modify the worksheet name by adding a numerical suffix
            'Note: In contrast to the way Excel does it [starting at " (2)"], I start at " (1)"
            i = 0
            Do
                Err.Clear
                i = i + 1
                'Test that we are not in a never-ending loop
                iLoopBailer = iLoopBailer + 1
                If iLoopBailer > 1000 Then
                    If Len(sExistingName) > 0 Then
                        fReturnSafeWorksheetName = sExistingName
                        Exit Function
                    Else
                        MsgBox "After 1000 iterations, no safe worksheet name could be found for the worksheet '" & sStartName & "'." _
                            , vbCritical + vbOKOnly, "Worksheet name failure"
                        
                        If Not bContinueOnFailure Then EndGracefully
                    End If
                End If
                'Determine the characters from the provided name to retain in the modified name
                'Recalculate this on each iteration to account for increases in the length of i as the iterations proceed...
                iNCharsToUseInNewName = Application.WorksheetFunction.Min(31, Len(sStartName) + Len(" ()") + Len(CStr(i))) - (Len(" ()") + Len(CStr(i)))
                fReturnSafeWorksheetName = Left(sStartName, iNCharsToUseInNewName) & " (" & CStr(i) & ")"
            
            'Test whether this name is in use
            Loop While fWorksheetNameIsInUse(sWorksheetName:=fReturnSafeWorksheetName, wb:=wb)
        Else
        'At this point, the worksheet name should be safe
            fReturnSafeWorksheetName = sStartName
        End If
        
    End If
End Function
Function fWorksheetNameIsInUse(ByRef sWorksheetName As String, Optional ByRef wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
        fWorksheetNameIsInUse = Len(wb.Worksheets(sWorksheetName).Name) > 0
    On Error GoTo 0
End Function

'Private Sub testfWorksheetNameIsInUse()
'Debug.Print fWorksheetNameIsInUse("Test Range", ThisWorkbook)
'Debug.Print fWorksheetNameIsInUse(LCase("Test Range"), ThisWorkbook)
'End Sub
'
'Private Sub testfReturnDefaultOpenWorksheetName()
'    fReturnDefaultOpenWorksheetName
'End Sub

Private Function fReturnDefaultOpenWorksheetName(Optional ByRef wb As Workbook)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Returns a workbook name in the series Sheet1, Sheet2, etc. which is not currently in the workbook '
' This can be used as a default when no valid sheet name can be found                               '
' Note: In contrast to the way Excel does it [starting at " (2)"], I start at " (1)"                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer, j As Integer
    If wb Is Nothing Then Set wb = ActiveWorkbook
    i = 0
    On Error Resume Next
        Do
            Err.Clear
            i = i + 1
            j = Len(wb.Worksheets("Sheet" & i).Name)
        Loop While Err.Number = 0
        Err.Clear
    On Error GoTo 0
    fReturnDefaultOpenWorksheetName = "Sheet" & i
End Function

Function fBreakPathForMsgBox(ByVal sPath As String) As String
'It appears a messagebox has a maximum width of 61 characters?, so this procedure breaks the file path more cleanly for better display
    
    Dim iEnd As Integer
    Dim iEndPrevious As Integer
    Dim iStart As Integer
    Dim iStartCur As Integer
    Dim iTotLen As Integer
    Dim iDoubleLen As Integer
    Dim iCounter As Integer
    Dim sPS As String
    Const iMaxLen As Integer = 61
    
    sPS = Application.PathSeparator
    iTotLen = Len(sPath)
    iDoubleLen = 2 * iTotLen
    iStart = 0
    iStartCur = -1
    iEnd = iMaxLen
    iEndPrevious = iEnd
    
    Do While iTotLen - iStart > iMaxLen And iTotLen < iDoubleLen
        'find the closest path separator to the break point
        iStart = InStrRev(sPath, sPS, iEnd, vbTextCompare)
        If iStart > 0 Then      'sPS found
            If iStart > iStartCur Then
            iStartCur = iStart
                'Add a line feed at that point
                sPath = Left(sPath, iStart) & vbLf & Right(sPath, iTotLen - iStart)
                'Increment for next iteration
                iTotLen = iTotLen + 1
                iStart = iStart + 1
                iEnd = iStart + iMaxLen
                If iEnd = iEndPrevious Then   'Prevent a recurring loop where multiple LFs are added
                    Exit Do
                Else
                    iEndPrevious = iEnd
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    
    fBreakPathForMsgBox = Replace(sPath, vbLf & vbLf, vbLf, , , vbTextCompare)
    
End Function
Sub testBreakString()
    MsgBox BreakString(fvReturnNameValue("APIKey"))
End Sub
Function BreakString(ByVal sIntial As String _
          , Optional ByVal iMaxLen As Integer = 61 _
          , Optional ByVal sToBreak As String = vbNullString) As String
'Based on the above, a more general purpose procedure to break strings for display
    
    Dim iStart As Integer       'These two set the range in the string within which to search for the path separator
    Dim iEnd As Integer         '
    Dim iTotLen As Integer      'Measures the length of the total string, and then increments larger as breaks are added
    Dim iDoubleLen As Integer   'Set to double the initial length of the string, as a limit to adding breaks
    
    BreakString = sIntial
    
    'Default sToBreak to path separator
    If sToBreak = vbNullString Then sToBreak = Application.PathSeparator
    
    iTotLen = Len(sIntial)
    iDoubleLen = 2 * iTotLen
    iStart = 0
    iEnd = iMaxLen
    
    Do While iTotLen - iStart > iMaxLen And iTotLen < iDoubleLen
        'find the closest path separator to the break point
        iStart = InStrRev(BreakString, sToBreak, iEnd)
        'Add a line feed at that point
        sIntial = Left(BreakString, iStart) & vbLf & Right(BreakString, iTotLen - iStart)
        'Increment for next iteration
        iTotLen = iTotLen + 1
        iStart = iStart + 1
        iEnd = iStart + iMaxLen
    Loop
    
End Function

Sub CopyRange(ByRef rngTarget As Range _
            , ByRef rngSource As Range _
   , Optional ByVal bIncludeNumberFormat As Boolean = True)
''''''''''''''''''''''''''''''''''''''''''''''''''
' Robustly copy one range to another range       '
' including the number formatting (if specified) '
' Note: This is only for small ranges,           '
'  and not preferred to Range.Copy Destination:= '
''''''''''''''''''''''''''''''''''''''''''''''''''
    If rngTarget.Rows.Count <> rngSource.Rows.Count _
    Or rngTarget.Columns.Count <> rngSource.Columns.Count Then
        MsgBox "The range to be copied is not the same size as the destination range." _
            , vbCritical + vbOKOnly, "Range mismatch"
        EndGracefully
    End If
    Dim i As Integer
    For i = 1 To rngTarget.Cells.Count
        rngTarget.Cells(i).Value = rngSource.Cells(i).Value
        If bIncludeNumberFormat Then rngTarget.Cells(i).NumberFormat = rngSource.Cells(i).NumberFormat
    Next i
End Sub

Function DoFreezePanes(ByRef wks As Worksheet _
                     , ByVal sRng As String _
                     , Optional ByRef wbk As Workbook) As Boolean
'Set FreezePanes in such a way that a possible error does not throw the whole process
'Value can be tested to see if an error occurred
    On Error Resume Next
    If Not wbk Is Nothing Then wbk.Activate
    wks.Activate
    wks.Range("A2").Select
    ActiveWindow.FreezePanes = True
    If Err.Number <> 0 Then
        Err.Clear
        DoFreezePanes = False
    Else
        DoFreezePanes = True
    End If
    On Error GoTo 0

End Function
Private Sub TestfGenerateSuggestedName()
fGenerateSuggestedName "W"
End Sub

Sub SuggestDataTarget()
'Used to auto-name a target file for extractions

    Dim sNewName As String
    Dim sWorksheetPrefix As String
    Dim bCheckWorksheet As Boolean
    
    bCheckWorksheet = (ActiveSheet.Name = Sheet3.Name Or ActiveSheet.Name = Sheet11.Name)
    If Not bCheckWorksheet Then EndGracefully
    
    sWorksheetPrefix = IIf(ActiveSheet.Name = Sheet3.Name, vbNullString, "W")
    
    'Check that no outstanding inputs are required (other than file name inputs)
    If fvReturnNameValue(sWorksheetPrefix & "ErrorDisplay1") = fvReturnNameValue(sWorksheetPrefix & "InputsCompleteMessage") _
        Or (fvReturnNameValue(sWorksheetPrefix & "ErrorDisplay1") = fvReturnNameValue(sWorksheetPrefix & "AllInputAreasMessage") _
            And fvReturnNameValue(sWorksheetPrefix & "ErrorDisplay2") = fvReturnNameValue(sWorksheetPrefix & "TargetFileMessage1")) _
        Or fvReturnNameValue(sWorksheetPrefix & "ErrorDisplay1") = fvReturnNameValue(sWorksheetPrefix & "TargetFileMessage1") _
        Or fvReturnNameValue(sWorksheetPrefix & "ErrorDisplay1") = fvReturnNameValue(sWorksheetPrefix & "TargetFileMessage2") _
    Then
        'All necessary inputs are complete, so a name can be suggested
        sNewName = fGenerateSuggestedName(sWorksheetPrefix:=sWorksheetPrefix)
        'Recalculate to update error messages
        ActiveSheet.Calculate

        If Len(sNewName) > 0 Then
            Range(sWorksheetPrefix & "DataTarget").Value = sNewName
        Else
            MsgBox "A suggested name could not be generated.", vbCritical + vbOKOnly, "Name suggestion failed"
        End If
    Else
        'Some inputs are still needed before a name can be suggested
        MsgBox "Please complete all inputs before generating a suggested name.", vbInformation + vbOKOnly, "Inputs outstanding"
        EndGracefully
    End If
    
End Sub
Function fReturnQueryListAsOneString(Optional ByVal sWorksheetPrefix As String = vbNullString) As String
    Dim i As Integer
    Dim iMaxTerms As Integer
    Dim sQueryString As String
    iMaxTerms = IIf(sWorksheetPrefix = "W", 5, 30)
    
    'Get the full Query term list
    For i = 1 To iMaxTerms
        If Len(fvReturnNameValue(sWorksheetPrefix & "SearchTerm" & Format(i, "00"), bCheckForActiveWorkbook:=False)) > 0 Then
            sQueryString = sQueryString & "(" & fvReturnNameValue(sWorksheetPrefix & "SearchTerm" & Format(i, "00")) & ")"
        End If
    Next i

    fReturnQueryListAsOneString = sQueryString

End Function
Function fGenerateSuggestedName(Optional ByVal sWorksheetPrefix As String = vbNullString)
    If sWorksheetPrefix <> vbNullString And sWorksheetPrefix <> "W" Then
        MsgBox "An incorrect worksheet prefix has been specified.", vbCritical + vbOKOnly, "sWorksheetPrefix"
        EndGracefully
    End If
    
    Dim sTmpName As String
    Dim sQueryTerms As String
    Dim sGeo As String
    Dim sPath As String
    Dim iLenPath As Integer
    Dim iLenQuery As Integer
    Dim iLenTmpName As Integer
    Dim iLengthRemaining As Integer
    
''''    Dim i As Integer
''''    Dim iMaxTerms As Integer
''''    iMaxTerms = IIf(sWorksheetPrefix = "W", 5, 30)
''''
''''    'Get the full Query term list
''''    For i = 1 To iMaxTerms
''''        If Len(fvReturnNameValue(sWorksheetPrefix & "SearchTerm" & Format(i, "00"))) > 0 Then
''''            sQueryTerms = sQueryTerms & "(" & fvReturnNameValue(sWorksheetPrefix & "SearchTerm" & Format(i, "00")) & ")"
''''        End If
''''    Next i
    sQueryTerms = fReturnQueryListAsOneString(sWorksheetPrefix:=sWorksheetPrefix)
    
    RemoveIllegalCharsFromFileName sFilename:=sQueryTerms _
                                 , sReplacementChar:="_" _
                                 , bRemoveAllAdditional:=True
'        bRemovePeriod = True
'        bRemoveTilde = True
'        bRemoveExclamation = True
'        bRemoveHash = True
'        bRemoveAt = True
'        bRemoveDollar = True
'        bRemovePound = True
'        bRemovePercent = True
'        bRemoveCaret = True
'        bRemoveAmpersand = True
'        bRemoveBraces = True
    
    'Get the date specification
    If sWorksheetPrefix = "W" Then
        sTmpName = Format(IIf(Len(fvReturnNameValue("WStartDate")) = 0, #1/1/2004#, fvReturnNameValue("WStartDate")), "yyyy-mm-dd") & "--" _
            & Format(IIf(Len(fvReturnNameValue("WEndDate")) = 0, Date, fvReturnNameValue("WEndDate")), "yyyy-mm-dd")
    Else
        sTmpName = Format(fvReturnNameValue("StartDate"), "yyyy-mm-dd") & "--" _
            & Format(fvReturnNameValue("EndDate"), "yyyy-mm-dd") _
            & "(" & IIf(fvReturnNameValue("DateResolution") = "Day", "Daily", fvReturnNameValue("DateResolution") & "ly") & ")"
    End If
    
    'Get the # samples
    If sWorksheetPrefix = vbNullString Then _
        sTmpName = sTmpName & ";" & fvReturnNameValue("Samples") & " sample" & IIf(fvReturnNameValue("Samples") = 1, vbNullString, "s")
    
    'Get the location
    If fvReturnNameValue(sWorksheetPrefix & "GeographicLevel") = "Worldwide" Then
        sGeo = "WW"
    Else
        sGeo = fvReturnNameValue(sWorksheetPrefix & "Descriptor")
    End If
    If sWorksheetPrefix = vbNullString And fvReturnNameValue("GeographicLevel") = "Region" And Len(fvReturnNameValue("Region")) = 0 Then
        sGeo = sGeo & "(ALL)"
    End If
''    ElseIf fvReturnNameValue(sWorksheetPrefix & "GeographicLevel") = "Country" _
''        Or sWorksheetPrefix = "W" _
''        Or Len(fvReturnNameValue("Region")) > 0 Then
''        sGeo = fvReturnNameValue(sWorksheetPrefix & "Descriptor")
''    ElseIf fvReturnNameValue(sWorksheetPrefix & "GeographicLevel") = "Region" Then
''        If Len(fvReturnNameValue("Region")) = 0 Then
''
''        Else
''
''        End If
''    End If
    sTmpName = sGeo & ";" & sTmpName
    
    'Get the function, domain and category when using the Google Trends Web interface
    If sWorksheetPrefix = "W" Then
        sTmpName = sTmpName & ";" & fvReturnNameValue("WFunction") _
        & ";" & fvReturnNameValue("WDomain")
        If Len(fvReturnNameValue("WCategory")) > 0 Then
            Dim sCat As String
            sCat = fvReturnNameValue("WCategory")
            RemoveIllegalCharsFromFileName sFilename:=sCat _
                                     , sReplacementChar:="_" _
                                     , bRemoveAllAdditional:=True
            sTmpName = sTmpName & ";" & sCat
        End If
    End If
        
    'Add the current date (of extraction)
    sTmpName = sTmpName & ";" & Format(Date, "yyyy-mm-dd")
    
    'Add the file extension
    sTmpName = sTmpName & ".xlsx"
    
    'Get the path
    sPath = IIf(Len(fReturnPathFromFileName(fvReturnNameValue(sWorksheetPrefix & "DataTarget"))) = 0 _
        , Application.DefaultFilePath _
        , fReturnPathFromFileName(fvReturnNameValue(sWorksheetPrefix & "DataTarget")))
        
    'Test the length of the various components
    iLenPath = Len(sPath)
    iLenQuery = Len(sQueryTerms)
    iLenTmpName = Len(sTmpName)
    iLengthRemaining = fvReturnNameValue("MaxFilePathLength") - (iLenPath + iLenTmpName)
    
    If iLengthRemaining < 0 Then
        'The name is going to be too long regardless of the query length, so let the user adjust it
        MsgBox "The suggested file name is too long." & vbCrLf & "Please edit it manually to be less than " _
            & fvReturnNameValue("MaxFilePathLength") & " characters." _
            , vbInformation + vbOKOnly, "Suggested name too long"
    Else
        If iLenQuery > iLengthRemaining - 1 Then
            sQueryTerms = Left(sQueryTerms, iLengthRemaining - 2) & ")"
            If bShowCompletionMsgBoxes Then _
                MsgBox "The query string in the suggested file name has been shortened so that the full path is less than " _
                & fvReturnNameValue("MaxFilePathLength") & " characters." _
                , vbInformation + vbOKOnly, "Query string shortened"
        End If
    End If
    
    fGenerateSuggestedName = sPath & sQueryTerms & ";" & sTmpName
    
End Function

Sub RemoveIllegalCharsFromFileName(ByRef sFilename As String _
                        , Optional ByVal sReplacementChar As String = vbNullString _
                        , Optional ByVal bRemoveAllAdditional As Boolean = False _
                        , Optional ByVal bRemovePeriod As Boolean = False _
                        , Optional ByVal bRemoveDoublePeriod As Boolean = False _
                        , Optional ByVal bRemoveTilde As Boolean = False _
                        , Optional ByVal bRemoveExclamation As Boolean = False _
                        , Optional ByVal bRemoveHash As Boolean = False _
                        , Optional ByVal bRemoveAt As Boolean = False _
                        , Optional ByVal bRemoveDollar As Boolean = False _
                        , Optional ByVal bRemovePound As Boolean = False _
                        , Optional ByVal bRemovePercent As Boolean = False _
                        , Optional ByVal bRemoveCaret As Boolean = False _
                        , Optional ByVal bRemoveAmpersand As Boolean = False _
                        , Optional ByVal bRemoveBrackets As Boolean = False _
                        , Optional ByVal bRemoveBraces As Boolean = False)
    
'https://docs.microsoft.com/en-us/windows/desktop/FileIO/naming-a-file

    If bRemoveAllAdditional Then
        bRemovePeriod = True
        bRemoveDoublePeriod = True
        bRemoveTilde = True
        bRemoveExclamation = True
        bRemoveHash = True
        bRemoveAt = True
        bRemoveDollar = True
        bRemovePound = True
        bRemovePercent = True
        bRemoveCaret = True
        bRemoveAmpersand = True
        bRemoveBrackets = True
        bRemoveBraces = True
    End If
    
    sFilename = Replace(sFilename, "", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "<", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, ">", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, sColon, sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, sQuote, sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "/", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "\", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "|", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "?", sReplacementChar, , , vbTextCompare)
    sFilename = Replace(sFilename, "*", sReplacementChar, , , vbTextCompare)
    
    If bRemovePeriod Then
        sFilename = Replace(sFilename, ".", sReplacementChar, , , vbTextCompare)
    ElseIf bRemoveDoublePeriod Then
        Do While InStr(1, sFilename, "..", vbTextCompare)
            sFilename = Replace(sFilename, "..", sReplacementChar, , , vbTextCompare)
        Loop
    End If
    If bRemoveTilde Then sFilename = Replace(sFilename, "~", sReplacementChar, , , vbTextCompare)
    If bRemoveExclamation Then sFilename = Replace(sFilename, "!", sReplacementChar, , , vbTextCompare)
    If bRemoveAt Then sFilename = Replace(sFilename, "@", sReplacementChar, , , vbTextCompare)
    If bRemoveHash Then sFilename = Replace(sFilename, "#", sReplacementChar, , , vbTextCompare)
    If bRemoveDollar Then sFilename = Replace(sFilename, "$", sReplacementChar, , , vbTextCompare)
    If bRemovePound Then sFilename = Replace(sFilename, "£", sReplacementChar, , , vbTextCompare)
    If bRemovePercent Then sFilename = Replace(sFilename, "%", sReplacementChar, , , vbTextCompare)
    If bRemoveCaret Then sFilename = Replace(sFilename, "^", sReplacementChar, , , vbTextCompare)
    If bRemoveAmpersand Then sFilename = Replace(sFilename, "&", sReplacementChar, , , vbTextCompare)
    If bRemoveBraces Then
        sFilename = Replace(sFilename, "{", sReplacementChar, , , vbTextCompare)
        sFilename = Replace(sFilename, "}", sReplacementChar, , , vbTextCompare)
    End If
    If bRemoveBrackets Then
        sFilename = Replace(sFilename, "[", sReplacementChar, , , vbTextCompare)
        sFilename = Replace(sFilename, "]", sReplacementChar, , , vbTextCompare)
    End If
End Sub
Sub DialBackColumnWidths(Optional ByRef wks As Worksheet _
                       , Optional ByVal iMaxColWidth As Integer = 40 _
                       , Optional ByVal lLastCol As Long)
    
    Dim c As Long
    
    If wks Is Nothing Then _
        Set wks = ActiveSheet
    With wks
        If lLastCol = 0 Then _
            lLastCol = .UsedRange.Columns(.UsedRange.Columns.Count).Column
            
        For c = 1 To lLastCol
            If .Columns(c).ColumnWidth > iMaxColWidth Then _
                .Columns(c).ColumnWidth = iMaxColWidth
        Next c
    End With
End Sub
