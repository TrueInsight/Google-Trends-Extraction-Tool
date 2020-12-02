Attribute VB_Name = "mBlockPaste"
Option Explicit
Option Private Module
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This module modifies code posted by Ken Puls at http://www.vbaexpress.com/kb/getarticle.php?kb_id=373 '
' to prevent users from pasting into the workbook, as that would                                       '
' mess up the names and conditional formatting which it relies on.                                     '
' The code is called from the ThisWorkbook module (Open, BeforeClose, Activate, Deactivate,            '
' as well as at the end of the CompleteReporting sub and from mWorkbookSetup.                          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function fNextWorkbookIsInProtectedView() As Boolean
    On Error Resume Next
    Dim s As String
    s = ActiveWorkbook.Name
    If Err.Number <> 0 Then
        Err.Clear
        fNextWorkbookIsInProtectedView = True
    Else
        fNextWorkbookIsInProtectedView = False
    End If
End Function

Sub TurnOnPaste()
    ToggleCutCopyAndPaste True
    Application.StatusBar = vbNullString
End Sub

Sub ToggleCutCopyAndPaste(Allow As Boolean)
     'Activate/deactivate cut, copy, paste and pastespecial menu items
    On Error GoTo TCCP_Err
    
    Call EnableMenuItem(21, Allow) ' cut
    Call EnableMenuItem(19, Allow) ' copy
    Call EnableMenuItem(22, Allow) ' paste
    Call EnableMenuItem(755, Allow) ' pastespecial

     'Activate/deactivate drag and drop ability
    Application.CellDragAndDrop = Allow

     'Activate/deactivate cut, copy, paste and pastespecial shortcut keys
    With Application
        Select Case Allow
        Case Is = False
            .OnKey "^c", "CopyDisabled"
            .OnKey "^v", "PasteDisabled"
            .OnKey "^x", "CutDisabled"
            .OnKey "+{DEL}", "CutDisabled"
            .OnKey "^{INSERT}", "PasteDisabled"
        Case Is = True
            .OnKey "^c"
            .OnKey "^v"
            .OnKey "^x"
            .OnKey "+{DEL}"
            .OnKey "^{INSERT}"
        End Select
    End With

    Exit Sub
    
TCCP_Err:
    Beep
    Dim eN As Variant
    With Err
        eN = .Number
        .Clear
    End With
    On Error GoTo -1
    On Error Resume Next
    If eN <> 1004 Then
        Application.StatusBar = "Error " & eN & " was encountered while trying to turn " _
        & IIf(Allow, "on", "off") & " the cut/copy/paste functionality. Switch workbooks again to retry."
    Else
        MsgBox "Cut/copy/paste functionality cannot be turned " & IIf(Allow, "on", "off") _
        & " because the active workbook is protected.", vbCritical + vbOKOnly, "Google Trends Extraction tool: Cut/Copy/Paste"
    End If

End Sub

Sub EnableMenuItem(ctlId As Integer, Enabled As Boolean)
     'Activate/Deactivate specific menu item
     On Error GoTo EMI_Err
    Dim cBar As CommandBar
    Dim cBarCtrl As CommandBarControl
    For Each cBar In Application.CommandBars
        If cBar.Name <> "Clipboard" Then
            Set cBarCtrl = cBar.FindControl(ID:=ctlId, recursive:=True)
            If Not cBarCtrl Is Nothing Then cBarCtrl.Enabled = Enabled
        End If
    Next

    Exit Sub

EMI_Err:
    Beep
    Dim eN As Variant
    With Err
        eN = .Number
        .Clear
    End With
    On Error GoTo -1
    On Error Resume Next
    If eN <> 1004 Then
        Application.StatusBar = "Error " & eN & " was encountered while trying to turn " _
        & IIf(Enabled, "on", "off") & " the cut/copy/paste functionality. Switch workbooks again to retry."
    Else
        MsgBox "Cut/copy/paste functionality cannot be turned " & IIf(Enabled, "on", "off") _
        & " because the active workbook is protected.", vbCritical + vbOKOnly, "Google Trends Extraction tool: Cut/Copy/Paste"
    End If
'    Beep
'    With Err
'        If .Number <> 1004 Then
'            .Clear
'            On Error Resume Next
'            Application.StatusBar = "Error " & .Number & " was encountered while trying to turn " _
'            & IIf(Enabled, "on", "off") & " the cut/copy/paste functionality. Switch workbooks again to retry."
'        Else
'            .Clear
'            On Error Resume Next
'            MsgBox "Cut/copy/paste functionality cannot be turned " & IIf(Enabled, "on", "off") _
'            & " because the active workbook is protected.", vbCritical + vbOKOnly, "Cut/Copy/Paste"
'        End If
'        .Clear
'    End With

End Sub

Sub CutDisabled()
    CutCopyPasteDisabled "cut from"
End Sub
Sub CopyDisabled()
    CutCopyPasteDisabled "copy from"
End Sub
Sub PasteDisabled()
    CutCopyPasteDisabled "paste into"
End Sub
Sub CutCopyPasteDisabled(ByRef sFunc As String)
     'Inform user that the functions have been disabled
'    MsgBox "Sorry!  Cutting, copying and pasting have been disabled in this workbook!"
    Dim i
    For i = 1 To 3
        Beep
    Next i
    If Not fNextWorkbookIsInProtectedView Then _
        Application.StatusBar = Replace(sPasteDisallowMsg, "paste into", sFunc, , , vbTextCompare)
End Sub

