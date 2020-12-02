Attribute VB_Name = "mButtonCode"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub btnExtractData()
    bShowCompletionMsgBoxes = True
    Call mGoogleTrendsInfoExtraction.DrawSample
End Sub
Sub btnExtractDataWeb()
    bShowCompletionMsgBoxes = True
    Call mGTWebFunctions.GetGTWeb
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub btnClearSpecifications()
    Call mWorkbookSetup.ClearSpecifications
End Sub

Sub btnClearSpecificationsWeb()
    Call mWorkbookSetup.ClearSpecificationsWeb
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Note that these two buttons on the two worksheets call the same sub--
' The sub itself determines which sheet to write the data to based on the content of the selected file.
Sub btnSuggestDataTarget()
    Call mMultiPurposeProcedures.SuggestDataTarget
End Sub
Sub btnSuggestDataTargetWeb()
    Call mMultiPurposeProcedures.SuggestDataTarget
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Note that these two buttons on the two worksheets call the same sub--
' The sub itself determines which sheet to write the data to based on the content of the selected file.
Sub btnLoadSpecification()
    bShowCompletionMsgBoxes = True
    Call mFileReadingFunctions.ReadValuesFromFile
End Sub
Sub btnLoadSpecificationWeb()
    bShowCompletionMsgBoxes = True
    Call mFileReadingFunctions.ReadValuesFromFile
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub btnResetQueriesUsed()
    Call mWorkbookSetup.ResetQueriesUsed(Sheet3)
End Sub
Sub btnResetQueriesUsedWeb()
    Call mWorkbookSetup.ResetQueriesUsed(Sheet11)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub btnViewAccount()
    Call mFileReadingFunctions.ViewAccount
End Sub
Sub ViewAccountWeb()
    Call mFileReadingFunctions.ViewAccount
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Note that these two buttons on the two worksheets call the same sub--
' The sub itself determines which sheet to write the data to based on the content of the selected file.
Sub btnAbout()
    Call mWorkbookSetup.ShowAbout
End Sub
Sub btnAboutWeb()
    Call mWorkbookSetup.ShowAbout
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This button is on the About page, and uses the value stored on the sheet by the ShowAbout procedure to determine which worksheet to return to
Sub btnReturn()
    Call mWorkbookSetup.HideAbout
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub btnHelp()
    Call mFileReadingFunctions.OpenHelpFile
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
