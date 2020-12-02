Attribute VB_Name = "mWorkbookSetup"
'Google Trends Extended for Health Information Extraction Tool
'XXXXXXXXXXXXXXXXXXXXXXXXXX, Jacques Raubenheimer
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This module contains functions to set up the workbook for final distribution '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowCurrentVersionNumber()
'Since the version number is stored in mPublicDeclarations, this is a quick way to check the current settings
    Debug.Print "Main Version: " & iMainVersionNumber & "; Sub Version: " & iSubVersionNumber & "; Build Version: " & iBuildVersionNumber
End Sub
Private Sub DecrementBuildVersionNumber()
'When the CreateReleaseVersion procedure fails, the build version is often already incremented (see ShowCurrentVersionNumber above)
'This dials it back one, so that it is not necessary to jump between modules and do it manually each time.

    Dim sFind As String
    Dim sReplace As String
    Dim lCurLine As Long
    Dim oVBProject As Object
    Dim oVBComponent As Object
    Dim oVBE_Module As Object
    Dim bDidReplace As Boolean
    
    Set oVBProject = ThisWorkbook.VBProject
    
    Set oVBE_Module = oVBProject.VBComponents("mPublicDeclarations").CodeModule
    
    'Increment Build version
    'Stop
    sFind = sBuildString & iBuildVersionNumber
    sReplace = sBuildString & iBuildVersionNumber - 1
    bDidReplace = False
    
    With oVBE_Module
        For lCurLine = 1 To .CountOfLines
            If InStr(1, .Lines(lCurLine, 1), sFind, vbTextCompare) > 0 Then
                .ReplaceLine lCurLine, sReplace 'Replace(.Lines(lCurLine, 1), sFind, sReplace, , , vbTextCompare)
                bDidReplace = True
                Exit For
            End If
        Next lCurLine
    End With    'oVBE_Module
    
    If Not bDidReplace Then MsgBox "Build number could not be incremented!", vbInformation + vbOKOnly, "Cannot find Build number line"
    
End Sub

Private Sub CreateReleaseVersion()
    
    'N.B.! Programmatic access to the VBA model must be accepted in the Excel options
    
    'N.B. This procedure increments the build version number.
    '     If the Main version or Sub version number must be incremented,
    '     do so manually in mPublicDeclarations and set the Build version number to -1,
    '     so that it is incremented to 0 here.
    
    ThisWorkbook.Save
    
    StartSmoothly
    Dim sReleaseVersion As String
    Dim iCopyrightYear As Integer
    Dim sExportDir As String
    Dim sPS As String
    sPS = Application.PathSeparator
    
    Application.DisplayAlerts = False
    
    iCopyrightYear = Year(Now)
    sReleaseVersion = " v" & iMainVersionNumber & "." & iSubVersionNumber & "." & iBuildVersionNumber + 1
    
'    IncrementBuildVersion iBuildVersionNumber
'Because of crashes, I incorporated IncrementBuildVersion into UpdateCopyright
    'Modify VB Code
    SetToProductionVersion
    UpdateCopyright sReleaseVersion, iCopyrightYear

    Application.DisplayAlerts = False
    
    With ThisWorkbook
'        .Save
'        .SaveAs Filename:=Replace(ThisWorkbook.FullName, ".xlsm", Replace(sReleaseVersion, ".", "_", , , vbTextCompare) & ".xlsm", , , vbTextCompare) _
'              , AddToMRU:=False

        OpenUpWorkbook

        'Export the VBA code
        sExportDir = ExportModules(sReleaseVersion)
        If Len(sExportDir) > 0 Then Debug.Print "Export code modules success"
        
        Dim i As Integer
        For i = .Sheets.Count To 1 Step -1
            Select Case .Sheets(i).Name
            'Sheets to keep:
            Case Sheet2.Name _
               , Sheet3.Name _
               , Sheet4.Name _
               , Sheet7.Name _
               , Sheet8.Name _
               , Sheet11.Name _
               , Sheet12.Name _
               , Sheet15.Name _
               , Sheet16.Name _
               , Sheet17.Name _
               , Sheet18.Name _
               , Sheet19.Name _
               , Sheet21.Name _
               , Sheet23.Name
                'Do nothing
            Case Else
                Application.StatusBar = "Deleting sheet " & i & " [" & .Sheets(i).Name & "]"
                .Sheets(i).Delete
            End Select
        Next i
        
        Application.StatusBar = vbNullString
        
        ClearSpecifications False, False
        ClearSpecificationsWeb False, False
        
        Application.EnableEvents = False
        
        With Sheet11
            .Activate
            .Range("WQueriesThisSession").ClearContents
            .Range("A1").Select
        End With
        ScrollToA1 Sheet11
        
        With Sheet3
            .Activate
            .Range("QueriesThisSession").ClearContents
            .Range("APIKey").Value = vbNullString
            .Range("A1").Select
        End With
        ScrollToA1 Sheet3
        
        HideAbout
        
        SetUpWorkbook

''        .Save
        .SaveAs Filename:=Replace(ThisWorkbook.FullName, ".xlsm", Replace(sReleaseVersion, ".", "_", , , vbTextCompare) & ".xlsm", , , vbTextCompare) _
              , AddToMRU:=False
        .SaveCopyAs sExportDir & sPS & .Name
        
    End With
    
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
    
    EndGracefully
    
    MsgBox "Release version successfully created.", vbInformation + vbOKOnly, "Release version " & Replace(sReleaseVersion, ".", "_", , , vbTextCompare)
    
End Sub

Private Sub UpdateCopyright(ByVal sReLVer As String, ByRef iCopyrightYear As Integer)

    Dim sFind As String
    Dim sReplace As String
    Dim lCurLine As Long
    Dim oVBProject As Object
    Dim oVBComponent As Object
    Dim oVBE_Module As Object
    Const lLastLineToCheck As Long = 2  'Sets a limit to speed up the code module checking. No copyright notice is current expected after line 2 of any module
    Dim bDidReplace As Boolean
    
    
    Set oVBProject = ThisWorkbook.VBProject
    
    Set oVBE_Module = oVBProject.VBComponents("mPublicDeclarations").CodeModule
    
    'Increment Build version
    'Stop
    sFind = sBuildString & iBuildVersionNumber
    sReplace = sBuildString & iBuildVersionNumber + 1
    bDidReplace = False
    
    With oVBE_Module
        For lCurLine = 1 To .CountOfLines
            If InStr(1, .Lines(lCurLine, 1), sFind, vbTextCompare) > 0 Then
                .ReplaceLine lCurLine, sReplace 'Replace(.Lines(lCurLine, 1), sFind, sReplace, , , vbTextCompare)
                bDidReplace = True
                Exit For
            End If
        Next lCurLine
    End With    'oVBE_Module
    
    If Not bDidReplace Then MsgBox "Build number could not be incremented!", vbInformation + vbOKOnly, "Cannot find Build number line"
    
    
    sFind = "'Copyright (C) 20??, Jacques Raubenheimer"
    sReplace = "'Copyright (C) " & iCopyrightYear & ", Jacques Raubenheimer"
    
    For Each oVBComponent In oVBProject.VBComponents
    
        With oVBComponent.CodeModule
'            Debug.Print .Name & " || " & oVBComponent.Type & " || " & .CountOfLines & " lines"
            For lCurLine = 1 To .CountOfLines
                If .Lines(lCurLine, 1) Like sFind Then
                    .ReplaceLine lCurLine, sReplace 'Replace(.Lines(lCurLine, 1), sFind, sReplace, , , vbTextCompare)
'                    .DeleteLines lCurLine
'                    .InsertLines lCurLine, sReplace
                    Exit For
                End If
                If lCurLine > lLastLineToCheck Then Exit For
            Next lCurLine
        End With    'oVBComponent.CodeModule
    
    Next oVBComponent

    With Sheet7
        .Visible = xlSheetVisible
        .Activate
        .Range("VersionNumber").Value = "Version " & sReLVer
        .Range("CopyrightYear").Value = "©" & iCopyrightYear
    End With

    Sheet3.Activate

End Sub

Private Function ExportModules(ByVal sReLVer As String) As String
    'I wrote this using Chip Pearson's GetFileExtension function
    ' http://www.cpearson.com/excel/vbe.aspx
    'Changed to a function, which returns the directory to which the modules are exported
    
    Dim sPS As String               'Path separator
    Dim sTargetDir As String
    Dim sExtension As String
    Dim oVBProject As Object
    Dim oVBComponent As Object
    Dim oVBE_Module As Object
    Dim lCurLine As Long
    Dim bDidReplace As Boolean
    
    sPS = Application.PathSeparator
    
    Set oVBProject = ThisWorkbook.VBProject
    
    If Not FileDirCheck(FileOrDir:=Directory, FullPath:=ThisWorkbook.Path & sPS & "Code exports" & sPS) Then _
        MkDir ThisWorkbook.Path & sPS & "Code exports"

    sTargetDir = ThisWorkbook.Path & sPS & "Code exports" & sPS & "CodeModuleExport " & Format(Now, "yyyy-mm-dd hh-mm") & sReLVer
    
    MkDir sTargetDir
    
    For Each oVBComponent In oVBProject.VBComponents
        sExtension = GetFileExtension(VBComp:=oVBComponent)
'        oVBComponent.Export Filename:=sTargetDir & sPS & oVBComponent.Name & "." & sExtension
        oVBComponent.Export Filename:=sTargetDir & sPS & oVBComponent.Name & sExtension
    Next oVBComponent
    
    ExportModules = sTargetDir
    
End Function
Public Function GetFileExtension(VBComp As Object) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    ' http://www.cpearson.com/excel/vbe.aspx
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'These are late bound
    'The Types are listed here: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-property-vba-add-in-object-model
    
    Select Case VBComp.Type
        Case 2                          'vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case 100                        'vbext_ct_Document
            GetFileExtension = ".cls"
        Case 3                          'vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case 1                          'vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
        
End Function

Private Sub SetToProductionVersion()
    Dim sFind As String
    Dim sReplace As String
    Dim oVBE_Module As Object
    Dim lCurLine As Long
    Dim bDidReplace As Boolean
    
    Set oVBE_Module = ThisWorkbook.VBProject.VBComponents("mPublicDeclarations").CodeModule
    
    'Set to production version
    sFind = sProductionVersionString
    sReplace = Replace(sProductionVersionString, "False", "True", , , vbTextCompare)
    
    With oVBE_Module
        For lCurLine = 1 To .CountOfLines
            If InStr(1, .Lines(lCurLine, 1), sFind, vbTextCompare) > 0 Then
                .ReplaceLine lCurLine, Replace(.Lines(lCurLine, 1), sFind, sReplace, , , vbTextCompare)
                bDidReplace = True
            End If
        Next lCurLine
    End With    'oVBE_Module
    
    If Not bDidReplace Then MsgBox "bProductionVersion could not be set to True!", vbInformation + vbOKOnly, "Cannot find bProductionVersion line"
    
End Sub

Private Sub IncrementBuildVersion(ByVal iBuild As Integer)
    Dim sFind As String
    Dim sReplace As String
    Dim oVBE_Module As Object
    Dim lCurLine As Long
    Dim bDidReplace As Boolean
    
    Set oVBE_Module = ThisWorkbook.VBProject.VBComponents("mPublicDeclarations").CodeModule
    
    'Increment Build version
    'Stop
    sFind = sBuildString & iBuild
    sReplace = sBuildString & iBuild + 1
    bDidReplace = False
    
    With oVBE_Module
        For lCurLine = 1 To .CountOfLines
            If InStr(1, .Lines(lCurLine, 1), sFind, vbTextCompare) > 0 Then
                .ReplaceLine lCurLine, sReplace 'Replace(.Lines(lCurLine, 1), sFind, sReplace, , , vbTextCompare)
                bDidReplace = True
                Exit For
            End If
        Next lCurLine
    End With    'oVBE_Module
    
    If Not bDidReplace Then MsgBox "Build number could not be incremented!", vbInformation + vbOKOnly, "Cannot find Build number line"
    
End Sub

Private Sub ScrollToA1(ByRef wks As Worksheet, Optional bFirstActivate As Boolean = False)
    With wks
        If bFirstActivate Then
            .Parent.Activate
            .Activate
        End If
        With .UsedRange
            ActiveWindow.SmallScroll _
                Up:=.Rows(.Rows.Count).Row _
              , ToLeft:=.Columns(.Columns.Count).Column
        End With
    End With
End Sub

Sub DoSave()
    Application.StatusBar = "Saving this file..."
    ThisWorkbook.Save
    Application.StatusBar = vbNullString
End Sub

Private Sub ProtectUnprotectWorkbook(ByRef bProtect As Boolean)
    If bProtect Then
        ThisWorkbook.Protect Password:=fvReturnNameValue("Password"), Structure:=True, Windows:=True
    Else
        ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
    End If
End Sub

Sub OpenUpWorkbook()
    DoOpenSheet Sheet11
    DoOpenSheet Sheet3
    ProtectUnprotectWorkbook False
    'ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
    TurnOnPaste
End Sub

Private Sub DoOpenSheet(ByRef sh As Worksheet)
    With sh
        .Unprotect Password:=fvReturnNameValue("Password")
        ShowOrHideSheetTabs 1
        .Range("M1:XFD1").Columns.Hidden = False
    End With
End Sub

Sub SetUpWorkbook()
    DoProtectSheet sh:=Sheet11, sNameOfRangeToSelect:="WSearchTerm01", lLastVisibleCol:=12, lLastVisibleRow:=32
    DoProtectSheet sh:=Sheet3, sNameOfRangeToSelect:="SearchTerm01", lLastVisibleCol:=12, lLastVisibleRow:=32
    DoProtectSheet sh:=Sheet7, sNameOfRangeToSelect:="Return", lLastVisibleCol:=11, lLastVisibleRow:=26
    DoProtectSheet sh:=Sheet8, sNameOfRangeToSelect:="CategorySelectorLevel1", lLastVisibleCol:=3, lLastVisibleRow:=17
    DoProtectSheet sh:=Sheet12, lLastVisibleCol:=7, lLastVisibleRow:=1428, iAllowLockedCells:=xlNoRestrictions
    ShowOrHideSheetTabs 2
    
    ProtectUnprotectWorkbook True
    'ThisWorkbook.Protect Password:=fvReturnNameValue("Password"), Structure:=True, Windows:=True
End Sub

Private Sub DoProtectSheet(ByRef sh As Worksheet _
                , Optional ByVal sNameOfRangeToSelect As String = vbNullString _
                , Optional ByVal lLastVisibleCol As Long _
                , Optional ByVal lLastVisibleRow As Long _
                , Optional ByVal iAllowLockedCells As XlEnableSelection = xlUnlockedCells)
                    
    Dim pwd As String
    pwd = fvReturnNameValue("Password")
    
    If sNameOfRangeToSelect = vbNullString Then sNameOfRangeToSelect = "A1"
    
    With sh
        .Unprotect Password:=pwd
        
        .Activate
        .Range(sNameOfRangeToSelect).Select
        
        With ActiveWindow
'            .DisplayHorizontalScrollBar = False
'            .DisplayVerticalScrollBar = False
'            .DisplayWorkbookTabs = False
            .DisplayGridlines = False
            .DisplayHeadings = False
        End With
        
        If lLastVisibleCol > 0 Then _
            .Range(.Cells(1, lLastVisibleCol + 1), .Cells(1, .Columns.Count)).Columns.Hidden = True
        If lLastVisibleRow > 0 Then _
            .Range(.Cells(lLastVisibleRow + 1, 1), .Cells(.Rows.Count, 1)).Rows.Hidden = True
        
        ProtectSheet sh, pwd, iAllowLockedCells
        
        ThisWorkbook.Save
    End With
End Sub

Sub ProtectSheet(ByRef sh As Worksheet _
               , ByVal sPwd As String _
               , Optional ByVal iAllowLockedCells As XlEnableSelection = xlUnlockedCells)
    'UserInterfaceOnly is not saved with the workbook, but for the moment, I am OK with that (http://www.cpearson.com/excel/Protection.aspx)

    'Name                value       Description
    'xlNoRestrictions    0           Anything can be selected.
    'xlNoSelection       -4142       Nothing can be selected.
    'xlUnlockedCells     1           Only unlocked cells can be selected.
    
    sh.EnableSelection = iAllowLockedCells
    
      sh.Protect Password:=sPwd _
                , DrawingObjects:=True _
                , contents:=True _
                , Scenarios:=True _
                , UserInterfaceOnly:=True _
                , AllowDeletingColumns:=False _
                , AllowDeletingRows:=False _
                , AllowFiltering:=False _
                , AllowFormattingCells:=False _
                , AllowFormattingColumns:=False _
                , AllowFormattingRows:=False _
                , AllowInsertingColumns:=False _
                , AllowInsertingHyperlinks:=False _
                , AllowInsertingRows:=False _
                , AllowSorting:=False _
                , AllowUsingPivotTables:=False
End Sub

Sub ShowOrHideSheetTabs(Optional ByVal iCallingProcedureOnOff As Integer = 0)
    Dim bTurnOnOrOff As Boolean
    ProtectUnprotectWorkbook False
    With ActiveWindow
        If iCallingProcedureOnOff = 2 Then          'Turn off
            bTurnOnOrOff = False
        ElseIf iCallingProcedureOnOff = 1 Then      'Turn on
            bTurnOnOrOff = True
        ElseIf iCallingProcedureOnOff = 0 Then      'Toggle
            bTurnOnOrOff = Not .DisplayWorkbookTabs 'The state of the .DisplayWorkbookTabs is used to determine whether to turn on or off all display settings
        End If
        
        'If bTurnOnOrOff Then ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
        
'        .DisplayHorizontalScrollBar = bTurnOnOrOff
'        .DisplayVerticalScrollBar = bTurnOnOrOff
'        .DisplayWorkbookTabs = bTurnOnOrOff
        .DisplayGridlines = bTurnOnOrOff
        .DisplayHeadings = bTurnOnOrOff
    End With
    'Application.DisplayFormulaBar = bTurnOnOrOff
    Dim i As Integer
    With ThisWorkbook
        For i = 1 To .Worksheets.Count
            With .Worksheets(i)
                If .Name <> Sheet3.Name And .Name <> Sheet11.Name Then
                    If bTurnOnOrOff Then
                        .Visible = xlSheetVisible
                    Else
                        .Visible = xlSheetVeryHidden
                    End If
                End If
            End With
        Next i
    End With
    If bTurnOnOrOff Then ProtectUnprotectWorkbook True
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Save
End Sub

Sub ShowAbout()
    Dim iReturnVal As Integer
    'Set the value of the sheet to return to to either the "Google Trends Web" sheet, if that is the active sheet,
    ' or the "Google Trends Extended Health" in all other cases (as the default)
    If ActiveSheet.Name = Sheet11.Name Then
        iReturnVal = 11
    Else
        iReturnVal = 3
    End If
    ShowHideAbout bShow:=True _
                , iReturnTo:=iReturnVal
    
    Sheet7.Range("Return").Select
End Sub

Sub HideAbout()
    Dim vReturnVal As Variant
    'Read which sheet to return to
    vReturnVal = fvReturnNameValue("Return")
    If vReturnVal <> 11 And vReturnVal <> 3 Then vReturnVal = 3
    ShowHideAbout bShow:=False _
                , iReturnTo:=vReturnVal
End Sub

Sub ShowHideAbout(ByVal bShow As Boolean _
                , Optional ByVal iReturnTo As Integer)
    ProtectUnprotectWorkbook False
    If bShow Then
        With Sheet7
            .Visible = xlSheetVisible
            .Activate
            .Range("Return").Value = iReturnTo
        End With
    Else
        With Sheet7
            .Visible = xlSheetHidden
            .Range("Return").ClearContents
        End With
        'Return to Sheet 11 if specified, otherwise Sheet3 as the default
        If iReturnTo = 11 Then
            Sheet11.Activate
        Else
            Sheet3.Activate
        End If
    End If
    ProtectUnprotectWorkbook True
End Sub
Private Sub ClearAllSpecs()
    ClearSpecifications False, False
    ClearSpecificationsWeb False, False
End Sub

Sub ClearSpecifications(Optional ByVal bAskToClear As Boolean = True, Optional bDisableDisplay As Boolean = True)
    
    'Application.EnableEvents = False
    If bDisableDisplay Then StartSmoothly
    
    On Error Resume Next
    Dim bClear As Boolean
    
    With Sheet3
        .Range("SearchTermList").ClearContents
        If bAskToClear Then
            bClear = (MsgBox("Clear samples?", vbYesNo, "Samples") = vbYes)
        Else
            bClear = True
        End If
        If bClear Then .Range("Samples").Value = 1  'ClearContents
        .Range("DataTarget").Value = vbNullString
        .Range("DateResolution").Value = "Month"
        .Range("StartDate").Value = #1/1/2004#
        If Month(Now) = 1 Then
            .Range("EndDate").Value = DateSerial(Year(Now) - 1, 12, 31)
        Else
            .Range("EndDate").Value = DateSerial(Year(Now), Month(Now), Application.WorksheetFunction.EoMonth(Now, -1))
        End If
        .Range("StartDate", "EndDate").Style = "InputYearMonth"
    
        If bAskToClear Then
            bClear = (MsgBox("Clear location?", vbYesNo, "Location") = vbYes)
        Else
            bClear = True
        End If
        If bClear Then
            .Range("Region").ClearContents
            .Range("Country").ClearContents
            .Range("GeographicLevel").Value = "Worldwide"
        End If
    End With
    
    On Error GoTo 0
    
    'Application.EnableEvents = True
    If bDisableDisplay Then EndGracefully
    
End Sub

Sub ClearSpecificationsWeb(Optional ByVal bAskToClear As Boolean = True, Optional bDisableDisplay As Boolean = True)
    
    'Application.EnableEvents = False
    If bDisableDisplay Then StartSmoothly
    
    Dim bClear As Boolean
    On Error Resume Next
    With Sheet11
        .Range("WSearchTermList").ClearContents
        .Range("WDomain").Value = "Google Search"
        .Range("WDataTarget").Value = vbNullString
        .Range("WStartDate").Value = #1/1/2004#
        .Range("WEndDate").ClearContents
        .Range("WStartDate", "WEndDate").Style = "InputFullDate"
        .Range("WFunction").Value = "Graph"
        .Range("WCategory").ClearContents
        
        If bAskToClear Then
            bClear = (MsgBox("Clear location?", vbYesNo, "Loction") = vbYes)
        Else
            bClear = True
        End If
        If bClear Then
            .Range("WRegion").ClearContents
            .Range("WCountry").ClearContents
            .Range("WGeographicLevel").Value = "Worldwide"
        End If
        .Range("WSearchTerm01").Select
    End With
    
    On Error GoTo 0
    
    'Application.EnableEvents = True
    If bDisableDisplay Then EndGracefully
End Sub

Private Sub SetHelpHyperlinks()
    Dim rn As Range
    Dim cl As Range
    Dim wks() As Worksheet
    Dim i As Integer
    
    ReDim wks(1 To 2)
    Set wks(1) = Sheet3
    Set wks(2) = Sheet11
    
    For i = LBound(wks) To UBound(wks)
        Set rn = wks(i).UsedRange
        For Each cl In rn.Cells
            If cl.Style = "i" Then
                
'                cl.Hyperlinks.Delete
                cl.Hyperlinks.Add Anchor:=cl _
                                , Address:=vbNullString _
                                , SubAddress:=cl.Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlA1, External:=False) _
                                , ScreenTip:="Click for help" _
                                , TextToDisplay:="i"
                cl.Style = "i"
            End If
        Next cl
    Next i

End Sub

Private Sub ExtractHelpMessages()
'This procedure and the next extracts all the Data validation messages from the user interface so that they can be preserved and rebuilt if need be.
    Application.Calculation = xlCalculationManual
    Dim r As Long
    r = 1
    With Sheet21.Range("A1:E1")
        .Value = Array("Sheet", "Cell", "Title", "Input message", "Length")
        .Font.Bold = True
    End With
    
    ExtractHelpMessageOneSheet shtSource:=Sheet3 _
                             , shtTarget:=Sheet21 _
                             , lDocumentationRow:=r
    ExtractHelpMessageOneSheet shtSource:=Sheet11 _
                             , shtTarget:=Sheet21 _
                             , lDocumentationRow:=r
    
    Application.Calculation = xlCalculationAutomatic

End Sub
Private Sub ExtractHelpMessageOneSheet(ByRef shtSource As Worksheet _
                                     , ByRef shtTarget As Worksheet _
                                     , ByRef lDocumentationRow As Long)
    Dim rng As Range
    
    With shtTarget
        For Each rng In shtSource.UsedRange.Cells
            If rng.Style = "i" Then
                lDocumentationRow = lDocumentationRow + 1
                .Rows(lDocumentationRow).RowHeight = 15
                .Cells(lDocumentationRow, 1) = shtSource.Name
                .Cells(lDocumentationRow, 2) = rng.Address
                .Cells(lDocumentationRow, 3) = rng.Validation.InputTitle
                .Cells(lDocumentationRow, 4).WrapText = False
                .Cells(lDocumentationRow, 4) = rng.Validation.InputMessage
                .Cells(lDocumentationRow, 4).FormulaR1C1 = "=LEN(RC[-1])"
            End If
        Next rng
        .Columns.AutoFit
        'Check that no columns are too wide after the autofitting
        DialBackColumnWidths wks:=shtTarget
    End With
End Sub

Private Sub BuildHelpMessages()
'This procedure reads through the list of help messages stored on the source worksheet and writes the
    
    CreateHelpStyle
    
    Dim wksSource As Worksheet
    Set wksSource = Sheet21
    Dim wbk As Workbook
    Set wbk = ThisWorkbook
    
    Dim r As Long
    r = 2
    With wksSource
        Do
            If Not fWorksheetExists(.Cells(r, 1).Value) Then
                MsgBox "The worksheet '" & .Cells(r, 1).Value & "' in '" & .Name & "'!" & .Cells(r, 1).Address & " could not be found!" _
                    , vbInformation + vbOKOnly, "Cannot set message in row " & r
            Else
                With wbk.Worksheets(.Cells(r, 1).Value).Cells(.Cells(r, 2).Value)
                    .Style = "i"
                    .Locked = False
                    
                    If Len(.Cells(r, 2).Value) = 0 Then
                        MsgBox "There is no help message in cell " & .Cells(r, 4).Address & " of the worksheet '" & .Name & "!'", vbInformation + vbOKOnly, "No Help message"
                    
                    ElseIf Len(.Cells(r, 2).Value) > 255 Then
                        MsgBox "The help message in cell " & .Cells(r, 4).Address & " of the worksheet '" & .Name & "!' is longer than 255 characters." _
                            & vbCrLf & "Only the first 255 characters will be assigned", vbInformation + vbOKOnly, "Help message too long"
                    
                        With .Validation
                            .InputTitle = .Cells(r, 3).Value
                            .InputMessage = Left(.Cells(r, 4).Value, 255)
                    
                        End With
                    Else
                        With .Validation
                            .InputTitle = .Cells(r, 3).Value
                            .InputMessage = .Cells(r, 4).Value
                        End With
                    End If
                End With
            End If
            
            r = r + 1
        
        Loop While Len(.Cells(r, 1)) > 0
        
    End With
End Sub

Private Sub CreateHelpStyle()
'This procedure creates a style with the name "i" and sets the various attributes needed to make it a help style
'Possible to-do item: turn off protection

    Dim oStyle As Style
    On Error Resume Next
    If Not fStyleExists(sStyleName:="i", wbk:=ThisWorkbook) Then
        Set oStyle = ThisWorkbook.Styles.Add(Name:="i")
        If Err.Number <> 0 Then
            MsgBox "The Help Style 'i' could not be created", vbCritical + vbOKOnly, "Style cannot be created"
            Err.Clear
        End If
    Else
        Set oStyle = ThisWorkbook.Styles("i")
        If Err.Number <> 0 Then
            MsgBox "The Help Style 'i' could not be assigned", vbCritical + vbOKOnly, "Style cannot be set"
            Err.Clear
        End If
    End If
    On Error GoTo 0
    
    With oStyle
        .IncludeNumber = True
        .NumberFormat = "@"
        
        .IncludeFont = True
        With .Font
            .Name = "Webdings"
            .Size = 12
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .ThemeColor = 4
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        .IncludeAlignment = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        
        .IncludeBorder = True
        .Borders(xlLeft).LineStyle = xlNone
        .Borders(xlRight).LineStyle = xlNone
        .Borders(xlTop).LineStyle = xlNone
        .Borders(xlBottom).LineStyle = xlNone
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        .IncludePatterns = True
        With .Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .IncludeProtection = True
        .Locked = False
        .FormulaHidden = False
    
    End With
    
End Sub
Private Function fStyleExists(ByVal sStyleName As String _
                   , Optional ByRef wbk As Workbook) As Boolean
    
    If wbk Is Nothing Then Set wbk = ActiveWorkbook
    Dim i As Integer
    For i = 1 To wbk.Styles.Count
        If wbk.Styles(i).Name = sStyleName Then
            fStyleExists = True
            Exit Function
        End If
    Next i
    
End Function

Private Function fWorksheetExists(ByVal sWorksheetName As String _
                       , Optional ByRef wbk As Workbook) As Boolean

    If wbk Is Nothing Then Set wbk = ActiveWorkbook

    Dim i As Integer
    For i = 1 To wbk.Worksheets.Count
        If wbk.Worksheets(i).Name = sWorksheetName Then
            fWorksheetExists = True
            Exit Function
        End If
    Next i

End Function

Private Sub AddNames()
'Not completed. Possibly use this to self-repair the named ranges. See JKP's Name Manager Pick-Up function
    Dim r As Long
    r = 2
    Dim wksNames As Worksheet
    Set wksNames = Sheet15
    
    With wksNames
        Do
            ThisWorkbook.Names.Add Name:=.Cells(r, 1) _
                                 , RefersTo:=.Cells(r, 12) _
                                 , Visible:=.Cells(r, 3) _
                                 , RefersToLocal:=.Cells(r, 2)
            r = r + 1
        
        Loop While Len(.Cells(r, 1)) > 0
    End With
End Sub

Sub ResetQueriesUsed(ByRef wks As Worksheet)
    
    StartSmoothly
    
    Dim sWorksheetPrefix As String
'    Dim iNewQueryNo As Integer
    Dim bInputSuccess As Boolean
    Dim vInputBoxResult As Variant
    
    wks.Unprotect Password:=fvReturnNameValue("Password")
    
    If wks.Name = Sheet3.Name Then
        sWorksheetPrefix = vbNullString
    ElseIf wks.Name = Sheet11.Name Then
        sWorksheetPrefix = "W"
    End If
    
    Do While Not bInputSuccess
        On Error Resume Next
        vInputBoxResult = InputBox("Please give the new value for the number of queries used:" _
                    , "Reset Queries used", 0)
        If Err.Number <> 0 Then
            MsgBox "You must provide a valid numeric value!", vbInformation + vbOKOnly, "Invalid number"
            Err.Clear
        ElseIf Len(vInputBoxResult) = 0 Then
            'Cancelled
            EndGracefully
        ElseIf Not IsNumeric(vInputBoxResult) Then
            MsgBox "The input box entry must be numeric!", vbInformation + vbOKOnly, "Negative number not allowed"
        ElseIf CInt(vInputBoxResult) < 0 Then
            MsgBox "The number of queries used cannot be negative!", vbInformation + vbOKOnly, "Negative number not allowed"
        ElseIf CInt(vInputBoxResult) > fvReturnNameValue(sWorksheetPrefix & "MaxQueriesPerDay") Then
            MsgBox "The number of queries used cannot be more than the number of permitted queries per day!", vbInformation + vbOKOnly, "Queries usev value too high"
        Else
            bInputSuccess = True
        End If
    Loop
    
    Range(sWorksheetPrefix & "QueriesThisSession").Value = CInt(vInputBoxResult)
    
    ProtectSheet wks, fvReturnNameValue("Password")

    EndGracefully
End Sub

Function fCurrentVersionNumber() As String
    fCurrentVersionNumber = "v" & iMainVersionNumber & "." & iSubVersionNumber & "." & iBuildVersionNumber
End Function
