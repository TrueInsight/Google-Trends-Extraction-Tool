VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub FillCategories()
'I used this to automate the building of the categories worksheet
Dim r As Long
Dim c As Long
For r = 2 To 1428
    For c = 1 To Cells(r, 8).Value - 1
        If Len(Cells(r, c).Value) = 0 Then Cells(r, c).Value = Cells(r - 1, c).Value
    Next c
Next r
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Len(Target.Value) = 0 Then
        'Blank cell was clicked, do nothing
        Cancel = True
    Else
        Dim i As Integer
        Dim c As Long
        Dim r As Long
        With Target
            c = .Column
            r = .Row
        End With
        
        Application.EnableEvents = False
        With Sheet8
            'Clear the old selection
            For i = 6 To 1 Step -1
                .Range("CategorySelectorLevel" & i).ClearContents
            Next i
            'Load the new selection
            For i = 1 To c
                .Range("CategorySelectorLevel" & i).Value = Sheet12.Cells(r, i).Value
            Next i
            
            .Select
            .Range("CategorySelectorLevel" & i - 1).Select
        End With
        
        If bProductionVersion Then
            ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
            Sheet12.Visible = xlSheetVeryHidden
        End If
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Deactivate()
    Application.StatusBar = vbNullString
End Sub
