Attribute VB_Name = "mFileReadingFunctions"
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
Public Enum FoD
    File = 1
    Directory = 2
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This module contains a function to check whether a file or directory exists, '
' and some functions to read information from text files,                     '
' as well as functions used to load file paths to the interface.              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub OpenHelpFile()
    Dim sHelpfile As String
    sHelpfile = ThisWorkbook.Path & Application.PathSeparator & "Google Trends Data Extraction Tool Help File.pdf"
    If FileDirCheck(FileOrDir:=File _
                  , FullPath:=sHelpfile) Then
        ThisWorkbook.FollowHyperlink sHelpfile
    Else
        MsgBox "The help file '" & sHelpfile & "' was not found!", vbCritical + vbOKOnly, "Help file not found"
    End If
End Sub

Sub ViewAccount()
    Dim sAccountName As String
    Dim sAccountURLString As String
    Dim vKeyFileContents As Variant
    
    vKeyFileContents = Split(fReadWholeFile(fvReturnNameValue("APIKey")), vbCrLf, , vbTextCompare)
    sAccountURLString = sAccountURL & vKeyFileContents(UBound(vKeyFileContents))
    On Error Resume Next
    ThisWorkbook.FollowHyperlink sAccountURLString
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number _
            & vbCrLf & Err.Description _
            , vbCritical + vbOKOnly _
            , "Cannot open account"
        Err.Clear
    End If
End Sub

Sub GetAPILocation(ByRef sFileLoc As String)
'This sub uses the FileDialog to allow the user to specify the location of the file containing the API key
'The argument is set ByRef so that it returns the file location string

    'This function should fire only when used from Sheet3 of this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Or Not ActiveSheet.Name = Sheet3.Name Then Exit Sub
    
    Dim dlgFindFiles As FileDialog
    'Dim sFilePath As String
    Set dlgFindFiles = Application.FileDialog(msoFileDialogFilePicker)
    With dlgFindFiles
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Text", "*.txt", 1
        .Filters.Add "Key", "*.key", 2
        .FilterIndex = 1
        .ButtonName = "Open"
        .Title = "Select API key file"
        If .Show = -1 Then
            sFileLoc = .SelectedItems(1)
        Else
            sFileLoc = vbNullString
        End If
    End With
    Set dlgFindFiles = Nothing
End Sub

Sub GetDataTargetLocation(ByRef sFileLoc As String)
'This sub uses the FileDialog to allow the user to specify the name and location of the file which will contain the extraction output
'The argument is set ByRef so that it returns the file location string
    
    'This function should fire only when used from Sheet3 or Sheet11 of this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Or _
        Not (ActiveSheet.Name = Sheet3.Name Or ActiveSheet.Name = Sheet11.Name) Then Exit Sub
    
    Dim dlgSaveFile As FileDialog
    Set dlgSaveFile = Application.FileDialog(msoFileDialogSaveAs)
    With dlgSaveFile
        .AllowMultiSelect = False
        '.Filters.Add "Excel workbook (*.xlsx)", "*.xlsx", 1
        .FilterIndex = 1
        .ButtonName = "Save"
        .Title = "Save the extracted data"
        If .Show = -1 Then
            sFileLoc = .SelectedItems(1)
'            For Each varSelectedFiles In .SelectedItems     'there is only one, so the loop counting is excluded
'                strDestinationFile = varSelectedFiles
'                Debug.Print strDestinationFile
'            Next varSelectedFiles
        Else
            sFileLoc = vbNullString
        End If
    End With
    Set dlgSaveFile = Nothing
End Sub

Function FileDirCheck(ByVal FileOrDir As Integer _
                    , ByRef FullPath As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This function is called from code, as well as from functions in the worksheets '
' to test the existence of files or file paths                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If FileOrDir = File Then
        FileOrDir = vbNormal
    ElseIf FileOrDir = Directory Then
        FileOrDir = vbDirectory
        'Extract only the path
        FullPath = Left(FullPath, InStrRev(FullPath, Application.PathSeparator, , vbTextCompare))
    End If
    
    FileDirCheck = Len(Dir(FullPath, FileOrDir)) > 0
    'Debug.Print FileDirCheck
End Function
Private Sub testPath()
Debug.Print fReturnPathFromFileName(fvReturnNameValue("DataTarget"))
End Sub
Function fReturnPathFromFileName(ByVal sFullPath As String _
                     , Optional ByVal bReportErr As Boolean = False _
                     , Optional ByVal bTestPath As Boolean = False) As String
    If Len(sFullPath) = 0 Then
        If bReportErr Then _
            MsgBox "No string was provided from which to extract the file directory!" _
                , vbCritical + vbOKOnly, "Zero-length file path string"
        Exit Function
    End If
    Dim iLastPS As Integer
    Dim sPS As String
    sPS = Application.PathSeparator
    
    iLastPS = InStrRev(sFullPath, sPS, , vbTextCompare)
    If iLastPS = 0 Then
        If bReportErr Then _
            MsgBox "No path separator was found in the file path: " _
            & vbCrLf & sFullPath _
                , vbCritical + vbOKOnly, "No path separator"
        Exit Function
    Else
        fReturnPathFromFileName = Left(sFullPath, iLastPS)
    End If
    If bTestPath Then
        If Not FileDirCheck(FileOrDir:=Directory _
                          , FullPath:=fReturnPathFromFileName) Then
            If bReportErr Then _
            MsgBox "The file path" & vbCrLf & fReturnPathFromFileName & vbCrLf _
                & "extracted from the full file path" & vbCrLf & sFullPath & vbCrLf _
                & "does not exist!" _
                , vbCritical + vbOKOnly, "File path does not exist"
                    
            fReturnPathFromFileName = vbNullString
        End If
    End If
    
End Function

Function fGetValueFromFile(sFileLoc As String)
'Currently, only used to read the API key from the file in which it is stored
    Dim iFile As Integer
    iFile = FreeFile
    Open sFileLoc For Input As #iFile
    Line Input #iFile, fGetValueFromFile    'Only read the API key here, in the first line
    Close #iFile
End Function

Function fReadWholeFile(sFile As String) As String
'Not currently used
    Dim iFile As Integer
    iFile = FreeFile
    On Error Resume Next
    Open sFile For Input As #iFile
    If Err.Number <> 0 Then
        MsgBox "An error ocurred while trying to read the file '" & sFile & "'.", vbCritical + vbOKOnly, "Cannot read file"
        fReadWholeFile = vbNullString
    Else
        fReadWholeFile = Input$(LOF(iFile), iFile)
    End If
    On Error GoTo 0
    
End Function

Function fWriteToFile(sFile As String _
                    , sFileContents As String _
                    , Optional bAppend As Boolean = False) As Boolean
'Write the data returned from an API call in a json string to a text file
'Returns true if no error was encountered

    'Check that the specified file does not already exist
    If Not bAppend And FileDirCheck(FileOrDir:=File _
                    , FullPath:=sFile) Then
        fWriteToFile = False
        Exit Function
    End If
    
    Dim iFile As Integer
    iFile = FreeFile
    On Error Resume Next
    If bAppend Then
        Open sFile For Append As #iFile
    Else
        Open sFile For Output As #iFile
    End If
    Print #iFile, sFileContents
    Close #iFile
    
    'Test to see if the process was error free
    If Err.Number <> 0 Then
        fWriteToFile = False
        Err.Clear
    Else
        fWriteToFile = True
    End If
End Function

Sub WriteQSvaluesToFile(ByRef wbkTarget As Workbook _
                      , ByRef wksSpecToStore As Worksheet)
'Write the full query specification information to an xlVeryHidden sheet
' in the output workbook, so that it can later be recalled by ReadValuesFromFile.
' This serves as an additional signature sheet to identify workbooks created by this tool.
    
    Dim wksTarget As Worksheet
''    Dim wbkSource As Workbook
    Dim wksRangeList As Worksheet
    Dim r As Long
    Dim rr As Long
    
    Set wksRangeList = Sheet23
    Set wksTarget = wbkTarget.Worksheets.Add
    With wksTarget
        .Name = "QS data"
        .Visible = xlSheetVeryHidden
        'Sign the workbook              :Row 1
        rr = rr + 1
        .Cells(rr, 1).Value = sSignaturePhrase
        'XX Add web link                :Row 2
        rr = rr + 1
        .Cells(rr, 1).Value = sWebLink
        rr = rr + 1
        'Add the calling worksheet name :Row 3
        .Cells(rr, 1).Value = "Extraction:"
        .Cells(rr, 2).Value = wksSpecToStore.Name
        'Add the date and time          :Row 4
        rr = rr + 1
        .Cells(rr, 1).Value = "Time stamp:"
        .Cells(rr, 2).Value = Now
        .Cells(rr, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        'Add the column headers (with a separator row) for the Query specification information
        rr = rr + 2                    ': Row 6
        .Range(.Cells(rr, 1), .Cells(rr, 3)) = Array("Range name", "RefersToLocal", "Value")
    End With
    
    ThisWorkbook.Activate   'Activate this workbook so that the names can be evaluated
    
    With wksRangeList
        r = 2
        Do While Len(.Cells(r, 1).Value) > 0
            If InStr(1, .Cells(r, 2).Value, wksSpecToStore.Name, vbTextCompare) Then
                rr = rr + 1
                wksTarget.Cells(rr, 1).Value = .Cells(r, 1).Value
                wksTarget.Cells(rr, 2).Value = .Cells(r, 2).Value
                wksTarget.Cells(rr, 3).Value = fvReturnNameValue(.Cells(r, 1).Value, False)
            End If
            r = r + 1
        Loop
    End With

    wksTarget.Range(wksTarget.Cells(3, 1), wksTarget.Cells(rr, 3)).Columns.AutoFit
    'Check that no columns are too wide after the autofitting
    DialBackColumnWidths wks:=wksTarget

End Sub

Sub ReadValuesFromFile()
'Open an old extraction output file, and copy the query specification from that file
' to this file for modification and re-extraction
    
    StartSmoothly

    Dim dlgFindFiles As FileDialog
    Dim sFilePath As String         'Store the file path of the user-selected file
    Dim wbkSource As Workbook       'The user-selected previous data extraction workbook from which the query details will be extracted
    Dim wksQueryInfo As Worksheet   'The xlVeryHidden sheet containing the raw query specification information
'    Dim wksQS As Worksheet          'The Query Specification sheet, to be used in a plan B attemt at reading the specification if wksQueryInfo cannot be found
    Const sQSD As String = "QS data"
    Const sQS As String = "Query Specification"
    Dim sWorksheetQueryLoaded As String
    
    Set dlgFindFiles = Application.FileDialog(msoFileDialogFilePicker)
    With dlgFindFiles
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx", 1
        .FilterIndex = 1
        .ButtonName = "Open"
        .Title = "Select previous Google Trends extraction file"
        If .Show = -1 Then
            sFilePath = .SelectedItems(1)
        Else
            'The user cancelled
            sFilePath = vbNullString
            EndGracefully
        End If
    End With
    Set dlgFindFiles = Nothing
    
    On Error Resume Next
    Set wbkSource = Application.Workbooks.Open(Filename:=sFilePath, AddToMRU:=False)
    If Err.Number <> 0 Then
        MsgBox "There was an error reading the file '" & sFilePath & "'." _
            & vbCrLf & "The specifications could not be loaded." _
            , vbCritical + vbOKOnly, "File failed to load"
        EndGracefully
    End If

    'Use fWorksheetNameIsInUse to test if the worksheet exists
    If fWorksheetNameIsInUse(sWorksheetName:=sQSD, wb:=wbkSource) Then
        Set wksQueryInfo = wbkSource.Worksheets(sQSD)
        sWorksheetQueryLoaded = fsReadQueryInfoFromQSD(wksQueryInfo)
        
    ElseIf fWorksheetNameIsInUse(sWorksheetName:=sQS, wb:=wbkSource) Then
    'The QS data worksheet is not found, so try to find the Query specification worksheet
        Set wksQueryInfo = wbkSource.Worksheets(sQS)
        sWorksheetQueryLoaded = fsReadQueryInfoFromQS(wksQueryInfo)
    
    Else
        If MsgBox("The query information could not be extracted from the workbook you have selected." _
            & vbCrLf & "Do you want to keep the selected workbook open?", vbInformation + vbYesNo, "No query specification found") = vbNo Then _
            wbkSource.Close
        sWorksheetQueryLoaded = vbNullString
    
    End If
    
    If Len(sWorksheetQueryLoaded) + 0 Then
        wbkSource.Close SaveChanges:=False
        If bShowCompletionMsgBoxes Then _
        MsgBox "The query specfications were loaded to " & sWorksheetQueryLoaded & " from:" _
            & vbCrLf & fBreakPathForMsgBox(sFilePath), vbInformation + vbOKOnly, "Query loaded"
    End If
    
    'Update the styles if specifications were loaded to Sheet3
    If wksQueryInfo.Name = Sheet3.Name Then _
        Sheet3.Worksheet_Change Sheet3.Range("DateResolution")

    EndGracefully
    
End Sub

Function fsReadQueryInfoFromQSD(ByRef wks As Worksheet) As String
'Read the information from the QS data sheet and load it to the interface
'This procedure is launched from ReadValuesFromFile

    Dim sCallSheetName As String    'the name of the sheet to which the Query will be written
                                    ' ("Google Trends Web" or "Google Trends Extended Health")
    Dim sWholeAddress As String     'Used to check the whole address of the range to which the value should be written
    Dim sRangeName As String        'Used to indicate the range name to which the value should be written
    Dim r As Long
    
    With wks
        'Check that Extraction (i.e., calling sheet) is listed in row 3
        r = FindRowWithStringValue(wks:=wks _
                               , sValToFind:="extraction" _
                               , bPartialFind:=True _
                               , lRowExpected:=3 _
                               , lColToSearch:=1)
        If r > 0 Then
            sCallSheetName = .Cells(r, 2).Value
        Else
            r = FindRowWithStringValue(wks:=wks _
                                   , sValToFind:="extraction" _
                                   , bPartialFind:=True _
                                   , lRowExpected:=0 _
                                   , lColToSearch:=1)
            If r = 0 Then
                MsgBox "The name of the query specification source could not be identified." _
                    , vbCritical + vbOKOnly, "Source worksheet name not found"
                EndGracefully
            Else
                sCallSheetName = .Cells(r, 2).Value
            End If
        End If
        
        'Test that the retrieved sheet name exists in this workbook
        If Not fWorksheetNameIsInUse(sWorksheetName:=sCallSheetName, wb:=ThisWorkbook) Then
            If MsgBox("The name of the query specification source ('" & sCallSheetName & "') was retrieved, but could not be reconciled with this extraction tool." _
                & vbCrLf & "Do you want to keep the selected workbook open?" _
                , vbCritical + vbYesNo, "Source worksheet name invalid") = vbNo Then _
                wks.Parent.Close
            
            EndGracefully
        End If
        
''        If InStr(1, LCase(.Cells(3, 1).value), "extraction", vbTextCompare) > 0 Then
''            sCallSheetName = .Cells(3, 2).value
''        Else
''            r = 0
''            Do While r < .UsedRange.Rows(.UsedRange.Rows.Count).Row
''                r = r + 1
''                If InStr(1, LCase(.Cells(r, 1).value), "extraction", vbTextCompare) > 0 Then
''                    sCallSheetName = .Cells(r, 2).value
''                    Exit Do
''                End If
''            Loop
''        End If
''
''        If Len(sCallSheetName) = 0 Then
''        'Test that a sheet name was retrieved
''            MsgBox "The name of the query specification source could not be identified." _
''                , vbCritical + vbOKOnly, "Source worksheet name not found"
''            EndGracefully
''        ElseIf Len(sCallSheetName) = 0 Then
''            'Test that the retrieved sheet name exists in this workbook
''            On Error Resume Next
''            Dim s As String
''            s = ThisWorkbook.Worksheets(sCallSheetName).Cells(1, 1).value
''            If Err.Number <> 0 Then
''                MsgBox "The name of the query specification source ('" & sCallSheetName & "') was retrieved, but could not be reconciled with this extraction tool." _
''                , vbCritical + vbOKOnly, "Source worksheet name invalid"
''                EndGracefully
''            End If
''            On Error GoTo 0
''        End If
        
        'Now that the sheet name is identified, extract the components to the worksheet
        'Find the row with the header "Range name"
        r = FindRowWithStringValue(wks:=wks _
                               , sValToFind:="range name" _
                               , bPartialFind:=True _
                               , lRowExpected:=6 _
                               , lColToSearch:=1)
        If r = 0 Then
            r = FindRowWithStringValue(wks:=wks _
                                   , sValToFind:="range name" _
                                   , bPartialFind:=True _
                                   , lRowExpected:=0 _
                                   , lColToSearch:=1)
            If r = 0 Then
                MsgBox "The list of query data could not be found!", vbCritical + vbOKOnly, "Query data not found"
                EndGracefully
            End If
        End If
        
        'Activate the correct worksheet so that the range names can be used
        ThisWorkbook.Activate
        ThisWorkbook.Worksheets(sCallSheetName).Activate
        
        r = r + 1
        On Error Resume Next
        Do While Len(.Cells(r, 1).Value) <> 0 And r <= .UsedRange.Rows(.UsedRange.Rows.Count).Row
                        
            sWholeAddress = .Cells(r, 2).Value
            If Not InStr(1, sWholeAddress, sCallSheetName, vbTextCompare) > 0 Then
                MsgBox "The query specification data for the range '" _
                    & .Cells(r, 1).Value & "' at the address " & .Cells(r, 2).Value _
                    & " did not refer to the worksheet '" & sCallSheetName & "' as expected." _
                    , vbInformation + vbOKOnly, "Error in worksheet address"
            Else
                sRangeName = .Cells(r, 1).Value
                ThisWorkbook.Worksheets(sCallSheetName).Range(sRangeName).Value = .Cells(r, 3).Value
                'Possible to-do item: If loading the value to the range name fails,
                ' then split apart the address in column 2 to worksheet name and cell address,
                ' and attempt to load the value that way
                If Err.Number <> 0 Then
                    MsgBox "An error was encountered while loading the query specification data for the range '" _
                        & .Cells(r, 1).Value & "' at the address " & .Cells(r, 2).Value _
                        , vbInformation + vbOKOnly, "Error: " & .Cells(r, 1).Value
                    Err.Clear
                End If
                
                r = r + 1
            
            End If
        Loop
'        If Err.Number <> 0 Then
'            MsgBox "Some errors were encountered while loading the list of query specification data." _
'                & vbCrLf & "Please check that the query specification has been loaded correctly." _
'                , vbInformation + vbOKOnly, "Errors encountered while loading query specification"
'        End If
    End With
    
    fsReadQueryInfoFromQSD = sCallSheetName
    
End Function

Function fsReadQueryInfoFromQS(ByRef wks As Worksheet) As String
'If the raw data sheet cannot be found, then attempt to retrieve the various components from the formatted and structured Query Specification worksheet
'This will primarily be used with workbooks created before the code to write the raw QS data sheet existed
'Read the information from the QS data sheet and load it to the interface
'This procedure is launched from ReadValuesFromFile if ReadQueryInfoFromQSD is not used
    
    'Dim wksTarget As Worksheet  'worksheet to which the query specification will be loaded
    
    'Because the query summaries are so different, two different procedures read in the values
    If InStr(1, wks.Cells(1, 1).Value, "Web", vbTextCompare) Then
        ReadQSWeb wksSource:=wks, wksTarget:=Sheet11
        fsReadQueryInfoFromQS = Sheet11.Name
    Else
        ReadQSGTe wksSource:=wks, wksTarget:=Sheet3
        fsReadQueryInfoFromQS = Sheet3.Name
    End If

End Function

Sub ReadQSWeb(ByRef wksSource As Worksheet, ByRef wksTarget As Worksheet)
    
    Dim rng As Range
    
    With wksTarget
    
        .Parent.Activate
        .Activate
    
        'WDataTarget "'Google Trends Web'!$C$9"
        '.Range("WDataTarget").value = wksSource.Parent.FullName
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WDataTarget" _
                    , vToAdd:=wksSource.Parent.FullName
        
        'WStartDate  "'Google Trends Web'!$F$3"  Start date:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Start date:")
        '.Range("WStartDate").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WStartDate" _
                    , wksSource:=wksSource _
                    , sFind:="Start date:" _
                    , OffsetCol:=1
        
        'WEndDate    "'Google Trends Web'!$F$4"  End date:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="End date:")
        '.Range("WEndDate").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WEndDate" _
                    , wksSource:=wksSource _
                    , sFind:="End date:" _
                    , OffsetCol:=1
    
        'WGeographicLevel    "'Google Trends Web'!$I$3"  Geographic Level:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Geographic Level:")
        '.Range("WGeographicLevel").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WGeographicLevel" _
                    , wksSource:=wksSource _
                    , sFind:="Geographic Level:" _
                    , OffsetCol:=1
    
        'WCountry    "'Google Trends Web'!$I$4"  Country:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Country:")
        '.Range("WCountry").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WCountry" _
                    , wksSource:=wksSource _
                    , sFind:="Country:" _
                    , OffsetCol:=1
    
        'WRegion "'Google Trends Web'!$I$5"  Region:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Region:")
        '.Range("WRegion").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WRegion" _
                    , wksSource:=wksSource _
                    , sFind:="Region:" _
                    , OffsetCol:=1
    
        'WFunction   "'Google Trends Web'!$I$8"  Function:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Function:")
        '.Range("WFunction").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WFunction" _
                    , wksSource:=wksSource _
                    , sFind:="Function:" _
                    , OffsetCol:=1
    
        'WDomain "'Google Trends Web'!$I$10" Search domain:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Search domain:")
        '.Range("WDomain").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="WDomain" _
                    , wksSource:=wksSource _
                    , sFind:="Search domain:" _
                    , OffsetCol:=1
    
        'WCategory   "'Google Trends Web'!$I$12" Category:
        Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Category:")
        If rng Is Nothing Then
            MsgBox "Could not find the specification 'Category:' for the range 'WCategory'" _
                , vbInformation + vbOKOnly, "Specification not found"
        Else
            On Error Resume Next
            
            If rng.Offset(0, 1).Value = "All" Or Len(rng.Offset(0, 1).Value) = 0 Then
                .Range("WCategory").Value = vbNullString
            Else
                .Range("WCategory").Value = rng.Offset(0, 1).Value
            End If
        
            If Err.Number <> 0 Then
                Err.Clear
                MsgBox "Could not find the range 'WCategory' to which to load the specification 'Category:'" _
                    , vbInformation + vbOKOnly, "Range name not found"
            End If
            
            On Error GoTo 0
        End If
        
        'WSearchTerm01   "'Google Trends Web'!$C$3"  Query term
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Query term")
        'i = 1
        'Do While Len(rng.Offset(i, 0).value) > 0
        '    .Range("WSearchTerm0" & i).value = rng.Offset(i, 0).value
        '    i = i + 1
        'Loop
        Dim r As Long
        For r = 1 To 5
            AddRangeValue wksTarget:=wksTarget _
                        , sRngName:="WSearchTerm" & Format(r, "00") _
                        , wksSource:=wksSource _
                        , sFind:="Query term" _
                        , OffsetRow:=r
        Next r

    End With
        
End Sub

Sub ReadQSGTe(ByRef wksSource As Worksheet, ByRef wksTarget As Worksheet)
    wksTarget.Parent.Activate
    wksTarget.Activate

    Dim rng As Range
    Dim i As Integer
    
    With wksTarget
    
        .Parent.Activate
        .Activate
    
        'DataTarget "'Google Trends Extended Health'!$F$14"
        '.Range("DataTarget").value = wksSource.Parent.FullName
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="DataTarget" _
                    , vToAdd:=wksSource.Parent.FullName
        
        'DateResolution "'Google Trends Extended Health'!$F$3"
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Frequency:")
        '.Range("DateResolution").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="DateResolution" _
                    , wksSource:=wksSource _
                    , sFind:="Frequency:" _
                    , OffsetCol:=1
        
        'StartDate   "'Google Trends Extended Health'!$F$4"  Start date:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Start date:")
        '.Range("StartDate").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="StartDate" _
                    , wksSource:=wksSource _
                    , sFind:="Start date:" _
                    , OffsetCol:=1
    
        'EndDate "'Google Trends Extended Health'!$F$5"  End date:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="End date:")
        '.Range("EndDate").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="EndDate" _
                    , wksSource:=wksSource _
                    , sFind:="End date:" _
                    , OffsetCol:=1
    
        'GeographicLevel "'Google Trends Extended Health'!$I$3"  Geographic Level:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Geographic Level:")
        '.Range("GeographicLevel").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="GeographicLevel" _
                    , wksSource:=wksSource _
                    , sFind:="Geographic Level:" _
                    , OffsetCol:=1
    
        'Country "'Google Trends Extended Health'!$I$4"  Country:
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Country:")
        '.Range("Country").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="Country" _
                    , wksSource:=wksSource _
                    , sFind:="Country:" _
                    , OffsetCol:=1
    
        'Region  "'Google Trends Extended Health'!$I$5"  Region:
        Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Region:")
        If rng Is Nothing Then
            MsgBox "Could not find the specification 'Region:' for the range 'Region'" _
                , vbInformation + vbOKOnly, "Specification not found"
        Else
            On Error Resume Next
            
            If Len(rng.Offset(0, 2)) > 0 Then
            'More than one region is specified, which indicates a multi-region request
                .Range("Region").Value = vbNullString
            Else
                .Range("Region").Value = rng.Offset(0, 1).Value
            End If
            
            If Err.Number <> 0 Then
                Err.Clear
                MsgBox "Could not find the range 'Region' to which to load the specification 'Region:'" _
                    , vbInformation + vbOKOnly, "Range name not found"
            End If
            
            On Error GoTo 0
        End If
    
        'Samples "'Google Trends Extended Health'!$F$9"
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Samples:")
        '.Range("Samples").value = rng.Offset(0, 1).value
        AddRangeValue wksTarget:=wksTarget _
                    , sRngName:="Samples" _
                    , wksSource:=wksSource _
                    , sFind:="Samples:" _
                    , OffsetCol:=1
    
        'SearchTerm01    "'Google Trends Extended Health'!$C$3"  Query term
        'Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:="Query term")
        'i = 1
        'Do While Len(rng.Offset(i, 0).value) > 0
        '    .Range("SearchTerm0" & Format(i, "00")).value = rng.Offset(i, 0).value
        'Loop
        Dim r As Long
        For r = 1 To 30
            AddRangeValue wksTarget:=wksTarget _
                        , sRngName:="WSearchTerm" & Format(r, "00") _
                        , wksSource:=wksSource _
                        , sFind:="Query terms" _
                        , OffsetRow:=r
        Next r

    End With
    
End Sub

Sub AddRangeValue(ByRef wksTarget As Worksheet _
                , ByVal sRngName As String _
                , Optional ByRef wksSource As Worksheet _
                , Optional ByVal sFind As String _
                , Optional ByVal vToAdd As Variant _
                , Optional OffsetRow As Long = 0 _
                , Optional OffsetCol As Long = 0)
'Add a value to a named range (sRngName) on a specified worksheet (wksTarget):
' -Directly,  using vToAdd, in which case none of the other optional arguments must be specified
' -From a range in another worksheet (wksSource) which is identified by searching for sFind
'  In these cases, vToAdd must not be specified, and wksSource, and sFind must be specified.
'  OffsetRow and OffsetCol are optional for this second use case

    Dim rng As Range
    
    If IsMissing(vToAdd) And Len(sFind) = 0 Then
        'This should never be trigerred
        MsgBox "Either a value in vToAdd, or a string to find on the worksheet in sFind must be specified." _
            , vbCritical + vbOKOnly, "Programming call error"
        Exit Sub
    End If
    
    If (Not wksSource Is Nothing And Len(sFind) = 0) Or (wksSource Is Nothing And Len(sFind) > 0) Then
        'This should never be trigerred
        MsgBox "Both the worksheet in wksSource, and the string to find on the worksheet in sFind must be specified." _
            , vbCritical + vbOKOnly, "Programming call error"
        Exit Sub
    End If
    
    If IsMissing(vToAdd) Then
        Set rng = frFindFirstOnSheet(wks:=wksSource, sFindWhat:=sFind)
        If rng Is Nothing Then
            MsgBox "Could not find the specification '" & sFind & "' for the range '" & sRngName & "'" _
                , vbInformation + vbOKOnly, "Specification not found"
            Exit Sub
        Else
            vToAdd = rng.Offset(OffsetRow, OffsetCol).Value
        End If
    End If
        
    On Error Resume Next
    
    wksTarget.Range(sRngName).Value = vToAdd
    
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Could not find the range '" & sRngName & "' to which to load the specification '" & sFind & "'" _
            , vbInformation + vbOKOnly, "Range name not found"
    End If
    
    On Error GoTo 0

End Sub

Private Function FindRowWithStringValue(ByRef wks As Worksheet _
                           , ByVal sValToFind As String _
                           , Optional ByVal bPartialFind As Boolean = True _
                           , Optional ByVal lRowExpected As Long = 0 _
                           , Optional ByVal lColToSearch As Long = 1) As Long
'Find the row containing a trigger value in a certain column
    
    FindRowWithStringValue = 0
    Dim r As Long
    With wks
        If lRowExpected > 0 Then
            If bPartialFind Then
                If InStr(1, LCase(.Cells(lRowExpected, lColToSearch).Value), sValToFind, vbTextCompare) > 0 Then _
                    FindRowWithStringValue = lRowExpected
            Else
                If .Cells(lRowExpected, lColToSearch).Value = sValToFind Then _
                    FindRowWithStringValue = lRowExpected
            End If
        Else
            r = 0
            If bPartialFind Then
                Do While r <= .UsedRange.Rows(.UsedRange.Rows.Count).Row
                    r = r + 1
                    If InStr(1, LCase(.Cells(r, lColToSearch).Value), sValToFind, vbTextCompare) > 0 Then
                        FindRowWithStringValue = r
                        Exit Do
                    End If
                Loop
            Else
                Do While r <= .UsedRange.Rows(.UsedRange.Rows.Count).Row
                    r = r + 1
                    If .Cells(lRowExpected, lColToSearch).Value = sValToFind Then
                        FindRowWithStringValue = r
                        Exit Do
                    End If
                Loop
            End If
        End If
    End With
    
End Function

Function frFindFirstOnSheet(ByRef wks As Worksheet _
                    , ByVal sFindWhat As String) As Range
    
    Dim rFound As Range
    
    With wks
        Set rFound = .UsedRange.Find(what:=sFindWhat _
                                  , After:=.UsedRange.Cells(.UsedRange.Cells.Count))
    End With
    
    'If rFound Is Nothing Then Set frFindFirstOnSheet = Nothing
    Set frFindFirstOnSheet = rFound
    
End Function
Sub SaveWithErrorHandling(Optional ByVal iSaveOrSaveAs = 2 _
                        , Optional ByRef wbk As Workbook _
                        , Optional ByVal sFilePath As String _
                        , Optional ByVal bAddToMRU As Boolean = False)
    
    If wbk Is Nothing Then Set wbk = ActiveWorkbook
    
    On Error GoTo SaveErr
    If iSaveOrSaveAs = 1 Then
        wbk.Save
    ElseIf iSaveOrSaveAs = 2 Then
        If sFilePath = vbNullString Then sFilePath = Application.DefaultFilePath
        
        wbk.SaveAs Filename:=sFilePath _
                 , AddToMRU:=bAddToMRU
    End If

    Exit Sub

SaveErr:
    With Err
        If .Number <> 0 Then
            MsgBox "An error occurred while trying to save the workbook '" & wbk.Name _
                & "' to the file location '" & sFilePath & "'." _
                & vbCrLf & "Error " & .Number & "--" & .Description
            .Clear
        End If
    End With
    On Error GoTo 0
End Sub

