Attribute VB_Name = "mCharting"
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

Sub AddChartToOutput(ByRef wbkNew As Workbook _
                   , ByRef wksMainSummary As Worksheet _
                   , ByRef wksSummary() As Worksheet _
                   , ByRef iMultipleCount As Integer _
                   , ByRef iMaxRow As Integer _
                   , ByRef sChartTitle As String _
                   , ByRef iDoMultiple As BuildSheets _
                   , ByRef sMultipleColHeads() As String _
                   , ByRef vQueryList() As Variant)
                
    Dim ch As Chart                     'Both of these are used in building the chart of the data plots
    Dim rng As Range
    Dim rngHead As Range
    Dim i As Integer
    Dim c As Range
    
    'Do a save at this point, just in case something bombs out with adding the chart
    'wbkNew.Save
    SaveWithErrorHandling iSaveOrSaveAs:=1 _
                        , wbk:=wbkNew
        
    'Now add the chart
    With wksMainSummary
    
        Set ch = fAddChart(ct:=xlLine _
                         , wb:=wbkNew _
                         , sSub:=IIf(iMultipleCount = 1, "Plot for " & .Cells(1, 16).Value _
                            , "Plot of " & IIf(iDoMultiple = BuildSheetsByQueryTerm, "Query terms", _
                            IIf(iDoMultiple = BuildSheetsByRegion, "Regions", ""))) _
                         , RngXVal:=.Range(.Cells(2, 1), .Cells(iMaxRow, 1)) _
                         , sTitle:=sChartTitle _
                         , bAddLegend:=(iMultipleCount > 1) _
                         , bScaleValuesTo100:=False _
                         , sValAxisTitle:="Probability of search occurrence (x 10,000,000)" _
                         , bAddDataLabels:=False)
                         ', sCatAxisTitle:="Date range" _

        .Activate
        
        If iMultipleCount = 1 Then
            Set rng = .Range(.Cells(2, 7), .Cells(iMaxRow, 7))
            Set rngHead = .Cells(1, 7)
            AddSeries ch:=ch _
                    , iSeriesNo:=1 _
                    , rngNameAddress:=rngHead _
                    , rngData:=rng _
                    , RngXVal:=.Range(.Cells(2, 1) _
                    , .Cells(iMaxRow, 1)) _
                    , bSetColourToBlack:=True
'            sTitle = .Cells(1, 9).value
        Else
            If iDoMultiple = BuildSheetsByRegion Then
                .Range(.Cells(1, 2), .Cells(1, iMultipleCount + 1)).Value = sMultipleColHeads
            ElseIf iDoMultiple = BuildSheetsByQueryTerm Then
                .Range(.Cells(1, 2), .Cells(1, iMultipleCount + 1)).Value = vQueryList
            End If
            
            For i = 1 To iMultipleCount
'                Stop
'                .Range(.Cells(2, i + 1), .Cells(iMaxRow, i + 1)).FormulaR1C1 = "=IF('" & wksSummary(i).Name & "'!RC7=0," & sQuote & sQuote & ",'" & wksSummary(i).Name & "'!RC7)"
                Set rng = .Range(.Cells(2, i + 1), .Cells(iMaxRow, i + 1))
                Set rngHead = .Cells(1, i + 1)
                rng.FormulaR1C1 = "=IF('" & wksSummary(i).Name & "'!RC7=0," & sQuote & sQuote & ",'" & wksSummary(i).Name & "'!RC7)"
                If i = 1 Then
                    AddSeries ch:=ch _
                            , iSeriesNo:=i _
                            , rngNameAddress:=rngHead _
                            , rngData:=rng _
                            , RngXVal:=.Range(.Cells(2, 1) _
                            , .Cells(iMaxRow, 1)) _
                            , bSetColourToBlack:=False
                Else
                    AddSeries ch:=ch _
                            , iSeriesNo:=i _
                            , rngNameAddress:=rngHead _
                            , rngData:=rng _
                            , bSetColourToBlack:=False
                End If
            Next i
            
            Set rngHead = .Range(.Cells(1, 1), .Cells(1, i))
            rngHead.Font.Bold = True
            'rngHead.Columns.AutoFit
            For Each c In rngHead.Cells
                If c.ColumnWidth > 40 Then c.ColumnWidth = 40
            Next c
            
            DoFreezePanes wks:=wksMainSummary, sRng:="B2"
            
'            Set rng = .UsedRange
'            sTitle = "Plot of all queries"
        End If
        
'        Stop
'        BuildChart rngData:=rng, sSub:=sTitle, RngXVal:=.Range(.Cells(2, 1), .Cells(iMaxRow, 1))
    End With

End Sub

Function fAddChart(ByRef ct As XlChartType _
                 , ByRef wb As Workbook _
                 , ByRef sSub As String _
                 , ByRef RngXVal As Range _
                 , Optional ByRef sTitle As String _
                 , Optional ByVal bAddLegend As Boolean = False _
                 , Optional ByVal bScaleValuesTo100 As Boolean = False _
                 , Optional ByVal sValAxisTitle As String _
                 , Optional ByVal sCatAxisTitle As String _
                 , Optional ByVal bAddDataLabels As Boolean = False) As Chart
'Add a chart of the data extracted from Google Trends and summarised on the Summary (or other) worksheet
'wb is the destination workbook
'sSub is the name of the chart sheet (it will be checked to confirm that it is a safe name)
'STitle is the chart title (if none, then sSub is used)
'RngXVal is the range of values to be plotted in the chart
'BAddLegend indicates whether to add a legend or not

'    Dim ch As Chart
    'Check that the name for the chart sheet is safe
    sSub = fReturnSafeWorksheetName(sStartName:=sSub, wb:=wb)
    
    'Add the chart as a chart sheet directly
    Set fAddChart = wb.Charts.Add
    With fAddChart
        .Location where:=xlLocationAsNewSheet, Name:=sSub
        .Move After:=wb.Sheets(2)
        .FullSeriesCollection(1).XValues = RngXVal
        SetBlack .ChartArea.Format.TextFrame2.TextRange.Font.Fill
        .ChartType = ct
        If bAddLegend Then
            .SetElement msoElementLegendRightOverlay
        Else
            .SetElement msoElementLegendNone
        End If
        
        .SetElement msoElementPrimaryValueGridLinesNone
        '.SetElement msoElementChartTitleAboveChart
        .SetElement msoElementChartTitleCenteredOverlay
        
        '.ChartTitle.Text = IIf(Len(sTitle) > 0, sTitle, sSub)
        'http://www.vbaexpress.com/forum/showthread.php?23716-Max-length-of-a-Chart-Title
        'suggests the following 'workaround' for long titles, which will automatically truncate the length
        .ChartTitle.Format.TextFrame2.TextRange.Characters.Text = IIf(Len(sTitle) > 0, sTitle, sSub)
        
        'Set the axis titles
        If Len(sValAxisTitle) > 0 Then
            If ct = xlBarClustered Then
                .SetElement msoElementPrimaryValueAxisTitleHorizontal
            Else
                .SetElement msoElementPrimaryValueAxisTitleRotated
            End If
            .Axes(xlValue, xlPrimary).AxisTitle.Text = sValAxisTitle
        End If
        If Len(sCatAxisTitle) > 0 Then
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = sCatAxisTitle
        End If
        
        If bAddDataLabels Then .SetElement msoElementDataLabelOutSideEnd
        
        If bScaleValuesTo100 Then
            With .Axes(xlValue)
                .MaximumScale = 100
                .MinimumScale = 0
                .MajorUnit = 10
                .MinorUnit = 1
            End With
        End If
    End With
    On Error Resume Next
    Do
        fAddChart.FullSeriesCollection(1).Delete
    Loop While Err.Number = 0
    Err.Clear
    On Error GoTo 0
End Function

Sub BuildChart(ByRef rngData As Range, ByRef sSub As String, ByRef RngXVal As Range)
    Dim ch As Chart
    'Add the chart as a chart sheet directly
    rngData.Parent.Activate
    Set ch = rngData.Parent.Parent.Charts.Add
    
'    ActiveChart.FullSeriesCollection(1).Delete
'    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "='Means of all Region series'!$B$1"
'    ActiveChart.FullSeriesCollection(1).values = "='Means of all Region series'!$B$2:$B$927"
'    ActiveChart.FullSeriesCollection(1).XValues = "='Means of all Region series'!$A$2:$A$927"
    
    
    With ch
        .FullSeriesCollection(1).XValues = RngXVal
        .SetSourceData Source:=rngData, PlotBy:=xlColumns   'PlotBy:=xlRowCol.xlColumns
        .ChartType = xlLine
        .SetElement msoElementLegendNone
        .SetElement msoElementPrimaryValueGridLinesNone
        .SetElement msoElementChartTitleAboveChart
        .ChartTitle.Text = sSub
        '.ChartTitle.Caption = sSub
        .Location where:=xlLocationAsNewSheet, Name:=sSub
        '.HasTitle = True
        .FullSeriesCollection(1).Format.Line.Weight = 0.5   'was 1
        .FullSeriesCollection(1).XValues = RngXVal
        
        'Add some black formatting
        SetBlack .ChartArea.Format.TextFrame2.TextRange.Font.Fill
        SetBlack .FullSeriesCollection(1).Format.Fill
        SetBlack .FullSeriesCollection(1).Format.Line
    End With
End Sub

Sub AddSeries(ByRef ch As Chart _
            , ByRef iSeriesNo As Integer _
            , ByRef rngNameAddress As Range _
            , ByRef rngData As Range _
            , Optional RngXVal As Range _
            , Optional ByRef bSetColourToBlack = True)
    With ch
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(iSeriesNo)
            If Not RngXVal Is Nothing Then .XValues = RngXVal
            
            .Name = "='" & rngNameAddress.Parent.Name & "'!" & rngNameAddress.Address
            .values = rngData
            .Format.Line.Weight = 0.5
            If bSetColourToBlack Then
                With .Format
                    SetBlack .Fill
                    SetBlack .Line
                End With
            End If
        End With
    End With
End Sub

Sub SetBlack(ByRef o As Object)
    On Error Resume Next    'Not all of the properties apply to all the chart objects, to it is simpler to just pass over those that don't apply
    With o
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    On Error GoTo 0
End Sub

' Deprecated:
Function fAddLineChart(ByRef wb As Workbook _
                 , ByRef sSub As String _
                 , ByRef RngXVal As Range _
                 , Optional ByRef sTitle As String _
                 , Optional ByVal bAddLegend As Boolean = False _
                 , Optional ByVal bScaleVertTo100 As Boolean = False) As Chart
'Add a chart of the data extracted from Google Trends and summarised on the Summary (or other) worksheet
'wb is the destination workbook
'sSub is the name of the chart sheet (it will be checked to confirm that it is a safe name)
'STitle is the chart title (if none, then sSub is used)
'RngXVal is the range of values to be plotted in the chart
'BAddLegend indicates whether to add a legend or not

'    Dim ch As Chart
    'Check that the name for the chart sheet is safe
    sSub = fReturnSafeWorksheetName(sStartName:=sSub, wb:=wb)
    
    'Add the chart as a chart sheet directly
    Set fAddLineChart = wb.Charts.Add
    With fAddLineChart
        .Location where:=xlLocationAsNewSheet, Name:=sSub
        .Move After:=wb.Sheets(1)
        .FullSeriesCollection(1).XValues = RngXVal
        SetBlack .ChartArea.Format.TextFrame2.TextRange.Font.Fill
        .ChartType = xlLine
        If bAddLegend Then
            .SetElement msoElementLegendRightOverlay
        Else
            .SetElement msoElementLegendNone
        End If
        .SetElement msoElementPrimaryValueGridLinesNone
        .SetElement msoElementChartTitleAboveChart
        .ChartTitle.Text = IIf(Len(sTitle) > 0, sTitle, sSub)
        .SetElement msoElementPrimaryValueAxisTitleRotated
        '.Axes(xlValue, xlPrimary).AxisTitle.Text = "Proportion of searches (x 10,000,000)"
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Probability of search occurrence (x 10,000,000)" _
            & IIf(bScaleVertTo100, vbLf & "[Scaled to 100]", vbNullString)
        
        If bScaleVertTo100 Then
            With .Axes(xlValue)
                .MaximumScale = 100
                .MinimumScale = 0
                .MajorUnit = 10
                .MinorUnit = 1
            End With
        End If
    End With
    On Error Resume Next
    Do
        fAddLineChart.FullSeriesCollection(1).Delete
    Loop While Err.Number = 0
    Err.Clear
    On Error GoTo 0
End Function

' Deprecated:
Function fAddBarChart(ByRef wb As Workbook _
                 , ByRef sSub As String _
                 , ByRef RngXVal As Range _
                 , Optional ByRef sTitle As String _
                 , Optional ByVal bAddLegend As Boolean = False _
                 , Optional ByVal bScaleHorizTo100 As Boolean = True _
                 , Optional ByVal sVertAxisTitle As String) As Chart
'Add a chart of the data extracted from some Google Trends Web functions
'wb is the destination workbook
'sSub is the name of the chart sheet (it will be checked to confirm that it is a safe name)
'STitle is the chart title (if none, then sSub is used)
'RngXVal is the range of values to be plotted in the chart's X-axis
'BAddLegend indicates whether to add a legend or not

'    Dim ch As Chart
    'Check that the name for the chart sheet is safe
    sSub = fReturnSafeWorksheetName(sStartName:=sSub, wb:=wb)
    
    'Add the chart as a chart sheet directly
    Set fAddBarChart = wb.Charts.Add
    With fAddBarChart
        .Location where:=xlLocationAsNewSheet, Name:=sSub
        .Move After:=wb.Sheets(1)
        .FullSeriesCollection(1).XValues = RngXVal
        SetBlack .ChartArea.Format.TextFrame2.TextRange.Font.Fill
        .ChartType = xlBarClustered
        If bAddLegend Then
            .SetElement msoElementLegendRightOverlay
        Else
            .SetElement msoElementLegendNone
        End If
        .SetElement msoElementPrimaryValueGridLinesNone
        .SetElement msoElementChartTitleAboveChart
        .ChartTitle.Text = IIf(Len(sTitle) > 0, sTitle, sSub)
        .SetElement msoElementPrimaryValueAxisTitleRotated
        '.Axes(xlValue, xlPrimary).AxisTitle.Text = "Proportion of searches (x 10,000,000)"
        .Axes(xlValue, xlPrimary).AxisTitle.Text = sVertAxisTitle
        ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
        
        If bScaleHorizTo100 Then
            With .Axes(xlValue)
                .MaximumScale = 100
                .MinimumScale = 0
                .MajorUnit = 10
                .MinorUnit = 1
            End With
        End If
    End With
    On Error Resume Next
    Do
        fAddBarChart.FullSeriesCollection(1).Delete
    Loop While Err.Number = 0
    Err.Clear
    On Error GoTo 0
End Function
