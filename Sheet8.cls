VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
    
    If Target.Column = 3 Then
        If Target.Row >= 3 And Target.Row <= 7 Then
            Me.Range(Me.Cells(Target.Row + 1, 3), Me.Cells(8, 3)).ClearContents
'            Dim i As Integer
'            For i = 7 To Target.Row + 1 Step -1
'                Me.Cells(i, 2).value = vbNullString
'            Next i
        End If
    End If

    Application.EnableEvents = True

End Sub
Sub ClearCategories()
    Application.EnableEvents = False
    
    With Sheet8
        .Range("CategorySelectorLevels").ClearContents
        .Range("CategorySelectorLevel1").Value = "All categories: 0"
    End With
    With Application
        .StatusBar = vbNullString
        .EnableEvents = True
    End With
End Sub
Sub CancelCategorySelector()
    Sheet11.Activate
    If bProductionVersion Then Me.Visible = xlSheetVeryHidden
    ThisWorkbook.Protect Password:=fvReturnNameValue("Password")
    
End Sub
Sub UseCategory()
    
    Application.EnableEvents = False
    
    With Sheet11
        .Activate
        .Range("WCategory").Value = fvReturnNameValue("CategorySelectorSelection")
    End With
    
    If bProductionVersion Then Me.Visible = xlSheetVeryHidden
    ThisWorkbook.Protect Password:=fvReturnNameValue("Password")
    
    Application.EnableEvents = True
End Sub
Sub ViewAllCategories()
    
    ThisWorkbook.Unprotect Password:=fvReturnNameValue("Password")
    
    With Sheet12
        .Visible = xlSheetVisible
        .Activate
        Dim r As Long
        Dim c As Long
        Dim lFoundLevel As Long
        c = 1
        r = 2
        lFoundLevel = 0
        'Select the cell in the list of categories that contains the selection in the selector
        Do While Len(Me.Range("CategorySelectorLevel" & c).Value) > 0
            For r = r To .UsedRange.Rows(.UsedRange.Rows.Count).Row
                If c = 1 Then
                    If .Cells(r, c).Value = Me.Range("CategorySelectorLevel" & c).Value Then
                        lFoundLevel = r
                        Exit For
                    End If
                Else
                    'For subsequent levels, check that the range is still within the preceding level's values
                    If .Cells(r, c).Value = Me.Range("CategorySelectorLevel" & c).Value Then
                        lFoundLevel = r
                        Exit For
                    ElseIf .Cells(r, c - 1).Value <> Me.Range("CategorySelectorLevel" & c - 1).Value Then
                        'Could not find right category, so exit
                        Exit Do
                    End If
                End If
            Next r
            c = c + 1
        Loop
        If lFoundLevel > 0 Then
            .Cells(lFoundLevel, IIf(c = 1, c, c - 1)).Select
        Else
            .Range("A1").Select
        End If
    End With
    Application.StatusBar = "Double click on a cell to load the cagetory"
    
End Sub
