VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.EnableEvents = False
    ThisWorkbook.Worksheets(Sheet3.Name).Range("MacroTest").Value = False
    ThisWorkbook.Save
    
    WriteLogEntry LogEventClose
    
    'Re-enable cut, copy, paste
    Call mBlockPaste.ToggleCutCopyAndPaste(True)
    Application.EnableEvents = True
End Sub

Private Sub Workbook_Open()
    
    If Application.ActiveProtectedViewWindow Is Nothing Then
        Application.EnableEvents = False
        Sheet3.Unprotect Password:=fvReturnNameValue("Password")
        Range("QueriesThisSession").Value = 0
        If Range("MacroTest").Value <> True Then Range("MacroTest").Value = True
    
        Sheet11.Unprotect Password:=fvReturnNameValue("Password")
        Range("WQueriesThisSession").Value = 0
        
        WriteLogEntry LogEventOpen
        
        'Test for the API Key
        TestAPIKey
        
        'Disable cut, copy, paste
        Call mBlockPaste.ToggleCutCopyAndPaste(False)
        ProtectSheet Sheet3, fvReturnNameValue("Password")
        Application.EnableEvents = True
    End If
End Sub

Private Sub Workbook_Activate()
    
    If Application.ActiveProtectedViewWindow Is Nothing Then
        Application.StatusBar = vbNullString
        
        'Disable cut, copy, paste
        Call mBlockPaste.ToggleCutCopyAndPaste(False)
        
        If bProductionVersion Then
            Application.DisplayFormulaBar = False
            ShowOrHideSheetTabs 2
        End If
    End If
End Sub
 
Private Sub Workbook_Deactivate()
    'Re-enable cut, copy, paste if permitted
    'If Not mBlockPaste.fNextWorkbookIsInProtectedView Then
    If Application.ActiveProtectedViewWindow Is Nothing Then  'If Not ActiveWindow Is Nothing Then
        Call mBlockPaste.ToggleCutCopyAndPaste(True)
        Application.StatusBar = vbNullString
    End If
    Application.DisplayFormulaBar = True
End Sub

Private Sub TestAPIKey()
    If Not fvReturnNameValue("CompletedAPIKey") Then
        'There is no API key file listed. Search for one.
        SearchAndLoadAPIKey "No API Key file is specified"
    Else
        'An API Key is listed. Check that the file exists
        ''Len(Dir(fvReturnNameValue("APIKey"), vbNormal)) > 0
        On Error Resume Next
        If Len(Dir(Sheet3.Range("APIKey").Value)) = 0 Then
            'The file pointing to the API key does not exist. Search for one.
            SearchAndLoadAPIKey "The file specified for the API Key (" _
                & fBreakPathForMsgBox(Sheet3.Range("APIKey").Value) & ") cannot be found"
        End If
        If Err.Number <> 0 Then
            If Err.Number = 52 Then 'Bad file name or number
            'The drive might be inaccesible if it is a network drive
            'Thus the file cannot be tested at the moment
            'Exit and do nothing
                Err.Clear
            End If
        End If
    End If
    
End Sub
Private Sub SearchAndLoadAPIKey(Optional ByRef sPreamble As String = vbNullString)
    Dim sPath As String
    Dim sKeyFile As String
    Dim sPS As String
    
    sPS = Application.PathSeparator
    
    sPath = ThisWorkbook.Path & sPS
    sKeyFile = Dir(sPath & "*.key", vbNormal)
    With Sheet3
        If Len(sKeyFile) > 0 Then
            
            If MsgBox(sPreamble & ", but we have found this file:" _
                & vbCrLf & fBreakPathForMsgBox(sPath & sKeyFile) _
                & vbCrLf & "Do you want to load it to the interface?" _
                , vbQuestion + vbYesNo, "Load API Key file") = vbYes Then
                
                .Range("APIKey").Value = sPath & sKeyFile
                ThisWorkbook.Save
            End If
        Else    'No API key file is found
            .Range("APIKey").Value = .Range("APIKey").Value
            .Range("APIKeyMessage").Calculate
            MsgBox "The file specified for the API Key cannot be found.", vbCritical + vbOKOnly, "API Key not found"
        End If
    End With
End Sub
