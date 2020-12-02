VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YearMonthForm 
   Caption         =   "Select Y/M"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3510
   OleObjectBlob   =   "YearMonthForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YearMonthForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private sCallingAddress As String
Private iStartingYear As Integer
Private iMinYear As Integer
Private sStartingMonth As String
Private dtReturnDate As Date
Private bUseYearAndMonth As Boolean
Private bFormWasCancelled As Boolean

Private Sub cbCancel_Click()
    bFormWasCancelled = True
    Me.Hide
    
    'Unload Me
    
End Sub

Private Sub cbOK_Click()
'    On Error Resume Next
On Error GoTo 0
    Dim iSelectedMonth As Integer
    iSelectedMonth = IIf(cxMonth.ListIndex = -1, 1, cxMonth.ListIndex + 1)
    'If year and month is selected, and the EndDate is being set, then even though the year was set to be no earlier than the start year, it is still possible that the
    ' year-month combination for the end date precedes that of the start date
    If bUseYearAndMonth And sCallingAddress = "EndDate" Then
        If DateSerial(cxYear.Value, iSelectedMonth, 1) <= DateSerial(Year(fvReturnNameValue("StartDate")), Month(fvReturnNameValue("StartDate")), 1) Then
            MsgBox "the date of " & cxMonth.Value & " " & cxYear.Value & " is " _
            & IIf(DateSerial(cxYear.Value, iSelectedMonth, 1) < DateSerial(Year(fvReturnNameValue("StartDate")), Month(fvReturnNameValue("StartDate")), 1), " earlier than ", " equal to ") _
            & "the starting date of " & Format(fvReturnNameValue("StartDate"), "MMMM YYYY") & "." _
            & vbCrLf & "It needs to be later than the starting date." _
            , vbCritical + vbOKOnly, "Invalid end date"
            Exit Sub
        End If
    End If
    
    dtReturnDate = DateSerial(CInt(cxYear.Value), iSelectedMonth, 1)
    If Err.Number <> 0 Then
        Err.Clear
        dtReturnDate = 0
    End If
    
    bFormWasCancelled = False
    On Error GoTo 0
    
    Me.Hide
    'Unload Me
End Sub

Private Sub UserForm_Activate()
    With Me
        'Remove any years in the Year box which are below the MinYear
        Dim i As Integer
        i = 0
        For i = .cxYear.ListCount To 1 Step -1
            If .cxYear.List(i - 1) < iMinYear Then .cxYear.RemoveItem (i - 1)
        Next i
        
        .cxYear.Text = CStr(iStartingYear)
        If bUseYearAndMonth Then
            .Caption = "Select Year and Month"
            .cxMonth.Text = sStartingMonth
            .cxMonth.Visible = True
            .lbMonth.Visible = True
        Else
            .Caption = "Select Year"
            .cxMonth.ListIndex = 0
            .cxMonth.Visible = False
            .lbMonth.Visible = False
        End If
    End With
End Sub

Private Sub UserForm_Initialize()
    With Me
        .cxYear.List = ReturnYearsForGoogleTrends(True)
        .cxMonth.List = ReturnArrayOfMonths
        .StartUpPosition = 0
        .Top = (Application.UsableHeight / 2) - (Me.Height / 2)
        .Left = (Application.UsableWidth / 2) - (Me.Width / 2)
'        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
'        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
Property Let YearAndMonth(sBoth As Boolean)
    bUseYearAndMonth = sBoth
End Property
Property Let CallingCell(sAddress As String)
    sCallingAddress = sAddress
End Property
Property Let StartingYear(iYear As Integer)
    iStartingYear = iYear
End Property
Property Let MinYear(iYear As Integer)
    iMinYear = iYear
End Property
Property Let StartingMonth(sMonth As String)
    sStartingMonth = sMonth
End Property
Property Get SelectedDate() As Date
    SelectedDate = dtReturnDate
End Property
Property Get FormCancel() As Boolean
    FormCancel = bFormWasCancelled
End Property
