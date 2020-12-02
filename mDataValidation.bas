Attribute VB_Name = "mDataValidation"
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

' Programatically set the data validation for the Start and End Date cells
' No longer used directly, but can still be invoked
Sub SetDataValidationStartDW()
    With Range("StartDate").Validation
        .Delete
        .Add Type:=xlValidateDate, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=DATE(2004,1,1)", _
             Formula2:="=DATE(YEAR(NOW()),MONTH(NOW()),DAY(NOW()))-2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "Start Date"
        .InputMessage = "Enter a starting date between 1 January 2004 and today, in the format yyyy/mm/dd" _
             & vbCrLf & "Double click to use a date picker"
        .ShowError = True
        .ErrorTitle = "Out-of-range date"
        .ErrorMessage = "The date must be between 1 January 2004 and two days earlier than today!"
    End With
End Sub
Sub SetDataValidationEndDW()
    With Range("EndDate").Validation
        .Delete
        .Add Type:=xlValidateDate, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="=MAX(StartDate,DATE(2004,1,1))", _
             Formula2:="=DATE(YEAR(NOW()),MONTH(NOW()),DAY(NOW()))-2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "End Date"
''I decided to enfore explicit setting of the end date
'        .InputMessage = "Enter an ending date between the start date and two days before today, in the format yyyy/mm/dd" _
'             & vbCrLf & "Leave empty to search up to the present (two days before today)" _
'             & vbCrLf & "Double click to use a date picker"
        .InputMessage = "Enter an ending date between the start date and two days before today, in the format yyyy/mm/dd" _
             & vbCrLf & "Double click to use a date picker"
        .ShowError = True
        .ErrorTitle = "Inadmissable date"
'        .ErrorMessage = "The date must be on or after the start date, and on or earlier than two days before today (or blank)!"
        .ErrorMessage = "The date must be on or after the start date, and on or earlier than two days before today!"
    End With
End Sub
