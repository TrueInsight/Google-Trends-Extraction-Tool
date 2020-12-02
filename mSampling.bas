Attribute VB_Name = "mSampling"
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains code which implements the multiple sampling strategy I devised       '
' to get around the restriction caused by Google's caching of request data for 24h.         '
' The net result is that the code can draw multiple samples,                                '
' each sampled afresh (i.e., not a cached repeat),                                          '
' and stitch them back together to create a better sample estimate for the data extraction. '
' The strategy I devised, with an explanation of the formulas, will be published.           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub TestDrawSamples()
'    DrawSamples True
'End Sub

'Private Sub TestMatchSamplesWithPeriods()
'    MatchSamplesWithPeriods DrawSamples, ReturnPeriodsAsArray
'End Sub

Public Function MatchSamplesWithPeriods(ByRef vSamples As Variant _
                                      , ByRef vPeriods As Variant) As Variant
    
    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Function
    
    'Uses DrawSamples and ReturnPeriodsAsArray to create an array of start and end dates for each sampling
    'The dimensioning of the output array is thus 2 x the number of samplings
    Dim V() As Variant
    ReDim V(1 To 2, LBound(vSamples, 2) To UBound(vSamples, 2))
    
    Dim i As Integer
    For i = LBound(vSamples, 2) To UBound(vSamples, 2)
        V(1, i) = vPeriods(1, vSamples(2, i))
        V(2, i) = vPeriods(2, vSamples(3, i))
    Next i
    MatchSamplesWithPeriods = V
End Function

Public Function DrawSamples(Optional ByVal bDoReporting As Boolean, Optional ByVal iNper As Integer, Optional ByVal iNSam As Integer) As Variant
'This is the procedure that takes the number of periods and splits them up
' according to the optimal sampling I defined (X_min=(S×2)-1)
' (first sample is whole period, then break total time into two parts,
' with breakpoint working towards start/end from middle
'This array of periods is then later simply matched to the corresponding dates requested (not in this sub)
    
    'This function should fire only when used in this workbook
    If Not ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Function
    
    Dim iSampleCounter As Long   'Sample Index counter
''    Dim iNPer As Integer    'Stores the number of time periods to be sampled
''    Dim iNSam As Integer    'Stores the number of samples needed
    Dim lMaxCominbations As Long    'The binomial coefficient as the total number of possible contiguous samples given the number of periods
    'Dim rTableStart As Range        'Sets the starting point for the results of the sampling exercise
'    Dim arSamples() As Integer            'Sets the range for the listing of the periods
    Dim arSamples() As Variant            'Sets the range for the listing of the periods
    Dim arNSamplesPerTimePeriod() As Integer       'Counts how many samples there are per time period
    Dim iSmallestSample As Integer  'Checks to see whether all the periods eventually reach the desired number of samples
    Dim iBreakPoint As Integer    'Determines the BreakPoint for the sampling
    Dim iStart As Integer       'Extracts the start of each sample
    Dim iEnd As Integer         'Extracts the end of each sample
    Dim iSamplingPattern As Integer 'A value from 1 to 4 to indicate the pattern inherent in the sampling breakdowns
    Dim iSeedValue As Integer       'Used to calculate the sampling range form the breakpoint, dependent on which sampling round is being used
    
    'Get the number of periods and samples for the sampling calculation
    If iNper = 0 Then _
    iNper = fvReturnNameValue("Periods")
    If iNSam = 0 Then _
    iNSam = fvReturnNameValue("Samples")
    If iNper = 0 Or iNSam = 0 Then
        MsgBox "Cannot run when the periods or the samples are zero!", vbCritical + vbOKOnly, "Incorrect speficiation"
        EndGracefully
    ElseIf iNper = 1 Then
        'Skip this, as only a single sample is allowed to be drawn
'        MsgBox "No sampling can be done because there is only one time period!", vbCritical + vbOKOnly, "Single time period"
'        EndGracefully
    End If

    'If more samples are needed than periods, then the optimal sampling cannot be used, and some overlap will occur
    If iNSam > iNper Then
        MsgBox "It is not possible to extract enough combinations that will cover the start and end of the period." _
            & vbCrLf & "The sampling will stop when the number of samples equals the number of periods." _
            & vbCrLf & "You should consider reducing the number of samples or increasing the time resolution to get more time periods." _
            , vbInformation + vbOKOnly, "Samples exceed periods"
    End If

    'Caculate the largest number of possible combinations
    'lMaxCominbations = (iNPer * (iNPer + 1)) / 2   'Binomial coefficient as max is a gross overestimate
    lMaxCominbations = fvReturnNameValue("SamplingsRequired")
    
    'Build the sample table
    ReDim arSamples(1 To 4, 1 To 1) 'arSamples has four dimensions({Index;Start;End;Range}) by as many dimensions as there are samples
    ReDim arNSamplesPerTimePeriod(1 To iNper)   'Count haw many times each period is represented across all the samples
    
    'Set these values to the # of samples and 1, so that sampling increments inwards over successive rounds
    'iBreakPoint = iNPer \ 2 + IIf(bEvenPeriods, 0, 1)
    iBreakPoint = iNper \ 2

    'The first sample is the whole time period
    iSampleCounter = 1
    arSamples(1, iSampleCounter) = iSampleCounter         'Index (not really necessary, but makes life easier)
    arSamples(2, iSampleCounter) = 1         'Start
    arSamples(3, iSampleCounter) = iNper     'End
    arSamples(4, iSampleCounter) = iNper     'Range (not really necessary, but makes life easier)

    'Increment the minimum sample count
    AddSamplesToArray arNSamplesPerTimePeriod, arSamples(2, iSampleCounter), arSamples(3, iSampleCounter), iSmallestSample
    
    iSamplingPattern = 0
    'Now do the successive smaller samples
    Do While iSampleCounter <= lMaxCominbations _
            And iSmallestSample < iNSam
        iSampleCounter = iSampleCounter + 1
        If iSampleCounter = 2 And iBreakPoint = iNper - iBreakPoint Then
            'When an even number of periods is used, the first round can only iterate twice, not four times
            iSamplingPattern = 3
        Else
            'Increment to the next sampling configuration
            iSamplingPattern = iSamplingPattern + 1
        End If
        
        If iSamplingPattern = 5 Then    'Reset after every 4th interation
            iSamplingPattern = 1
            iBreakPoint = iBreakPoint - 1
        End If
        
        If iSamplingPattern = 1 Or iSamplingPattern = 4 Then
            'For the first and fourth iterations, use the breakpoint value
            iSeedValue = iBreakPoint
        Else
            'For the second and third iterations, use the mirror of the breakpoint value
            iSeedValue = iNper - iBreakPoint
        End If
        
        If iSampleCounter Mod 2 = 0 Then
            'For even samples, start sampling from the beginning
            iStart = 1
            iEnd = iSeedValue
        Else
            'For odd samples, sample towards the end
            iStart = iNper - iSeedValue + 1
            iEnd = iNper
        End If
        
        ReDim Preserve arSamples(1 To 4, 1 To iSampleCounter)
        arSamples(1, iSampleCounter) = iSampleCounter
        arSamples(2, iSampleCounter) = iStart
        arSamples(3, iSampleCounter) = iEnd
        arSamples(4, iSampleCounter) = arSamples(3, iSampleCounter) - arSamples(2, iSampleCounter) + 1
         'Increment the minimum sample count
        AddSamplesToArray arNSamplesPerTimePeriod, arSamples(2, iSampleCounter), arSamples(3, iSampleCounter), iSmallestSample
    Loop
    
    'Return the array to the function call
    DrawSamples = arSamples
    
End Function

Sub AddSamplesToArray(ByRef Arr() As Integer, ByVal iStart As Integer, ByVal iEnd As Integer, ByRef iSmallestSample As Integer)
'Takes the start- and end periods (simply using a numerical index for each period, not a date--the dates are matched in later)
' and then builds an array with as many rows as there are periods from start to end
'Both Start (>=1) and End (<=Max) will vary on each iteration
'It also uses iSmallestInThisRound to check what the minimum number of times each period (identified by its numerical index) has been sampled
    Dim i As Integer
    Dim iSmallestInThisRound As Integer
    For i = iStart To iEnd
        Arr(i) = Arr(i) + 1
    Next i
    iSmallestInThisRound = Arr(1)
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) < iSmallestInThisRound Then iSmallestInThisRound = Arr(i)
    Next i
    iSmallestSample = iSmallestInThisRound
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''All the code below is no longer used.
'' It was used in the developement of the sampling strategy employed for this tool
'Sub WriteSamplesToWorksheet(ByRef arS() As Integer, Optional ByVal ws As Worksheet)
''This procedure is no longer used. It allowed me to write the sampling strategy to a worksheet
'' so that I could visualize what was being done
'    If ws Is Nothing Then Set ws = ActiveSheet
'    Dim r As Long
'    Dim c As Long
'
'    With ws
'        Dim startRow As Long
'        'Get the first empty row
'        startRow = 1 + .UsedRange.Rows(.UsedRange.Rows.Count).Row
'
'        Dim StartC As Long
'        StartC = 2
'
'        For r = LBound(arS, 1) To UBound(arS, 1)
'            For c = LBound(arS, 2) To UBound(arS, 2)
'                .Cells(startRow + r, StartC + c).value = arS(r, c)
'            Next c
'        Next r
'        .Cells(startRow + 1, StartC).value = "Index"
'        .Cells(startRow + 2, StartC).value = "Start"
'        .Cells(startRow + 3, StartC).value = "End"
'        .Cells(startRow + 4, StartC).value = "Range"
'
'    End With    'ws
'
'End Sub
'
Sub TestOptimisation()
Dim r As Long   'Row counter
Dim c As Long   'Column counter
Dim iNper As Integer    'Stores the number of time periods to be sampled
Dim iNSam As Integer    'Stores the number of samples needed
Dim lMaxCominbations As Long    'The binomial coefficient as the total number of possible contiguous samples given the number of periods
Dim ws As Worksheet     'References the worksheet
Dim V As Variant        'Used to protect the specification range (below) when resetting the worksheet (clearcontents)
Dim rSpecifications As Range    'Stores the range (A1:B2) where the spefications for the sampling exercise are stored
Dim rTableStart As Range        'Sets the starting point for the results of the sampling exercise
Dim rPeriods As Range           'Sets the range for the listing of the periods
Dim rNSamplesPerTimePeriod      'Sets the range containing the formulas which tell us how many samples there are per time period
Dim iMax As Integer, iMin As Integer    'Used to determine the number of samples to draw in each round (incremented from iNSam and 1 respectively)
Dim bToggleMaxMin As Boolean    'Toggles between sampling from each end: When it is false, use Min, when it is true, use Max
Dim bEvenSamples As Boolean     'Determines whether an even or odd number of samples and periods are specified, for best optimisation
Dim rStart As Range     'Extracts the start of each sample
Dim rEnd As Range       'Extracts the end of each sample
'Dim sQuote As String    'Used to simply creating formulas using quotes
'sQuote = Chr$(34)
Dim sB1Formula As String    'Stores the formula for Cell B1

Set ws = Sheet14
With ws
    'Store the specifications range by setting it to a range variable and then passing that to a variant
    Set rSpecifications = .Range(.Cells(1, 1), .Cells(2, 2))
    V = rSpecifications
'    sB1Formula = .Cells(1, 2).Formula

    'Get the number of periods and samples for the sampling calculation
    iNper = .Cells(1, 2)
    iNSam = .Cells(2, 2)
    If iNper = 0 Or iNSam = 0 Then
        MsgBox "Cannot run when the periods or the samples are zero!", vbCritical + vbOKOnly, "Incorrect speficiation"
        EndGracefully
    ElseIf iNper = 1 Then
        MsgBox "No smapling can be done because there is only one time period!", vbCritical + vbOKOnly, "Single time period"
        EndGracefully
    End If

    'If more samples are needed than periods, then the optimal sampling cannot be used, and some overlap will occur
    If iNSam > iNper Then
'        MsgBox "It is not possible to extract enough combinations that will cover the start and end of the period perfectly." _
'            & vbCrLf & "The sampling will create more overlaps on the start and end tails than are needed." _
'            & vbCrLf & "You should consider reducing the number of samples or increasing the time resolution to get more time periods." _
'            , vbInformation + vbOKOnly, "Samples exceed periods"
        MsgBox "It is not possible to extract enough combinations that will cover the start and end of the period." _
            & vbCrLf & "The sampling will stop when the number of samples equals the number of periods" _
             & vbCrLf & "You should consider reducing the number of samples or increasing the time resolution to get more time periods." _
            , vbInformation + vbOKOnly, "Samples exceed periods"
    End If

    'Caculate the largest number of possible combinations
'    lMaxCominbations = (iNPer * (iNPer + 1)) / 2
    lMaxCominbations = fvReturnNameValue("BinomialCoefficientForSampleMax")
    If lMaxCominbations > .Columns.Count - 2 Then
        lMaxCominbations = .Columns.Count - 2
'        MsgBox "The number of periods is too large to plot on this spreadsheet!", vbCritical + vbOKOnly, "Insufficient space"
'        End
    End If

    'Check whether the periods and samples are in agreement (both odd or both even)
    bEvenSamples = (iNSam Mod 2 = iNper Mod 2)

    'Wipe the sheet clean
    '.Cells.ClearContents
    .Cells.Clear

    'Replace A1:B2's values
    .Cells(1, 1).Resize(2, 2) = V
'    .Range(.Cells(1, 1), .Cells(2, 2)) = v
'    .Cells(1, 1) = "Periods:"
'    .Cells(2, 1) = "Samples:"
    'Replace B1's formula
'    If sB1Formula = vbNullString Or Left$(sB1Formula, 1) <> "=" Then _
'        sB1Formula = "=IFERROR(IF(OR(DateResolution=" & sQuote & "Day" & sQuote & _
'                     ",DateResolution=" & sQuote & "Month" & sQuote & _
'                     ",DateResolution=" & sQuote & "Year" & sQuote & _
'                     "),DATEDIF(StartDate,EndDate,LEFT(DateResolution,1))+1,IF(DateResolution=" & sQuote & _
'                     "Week" & sQuote & ",ROUNDUP((EndDate-StartDate+1)/7,0)," & sQuote & sQuote & "))," & sQuote & sQuote & ")"
'    .Cells(1, 2).Formula = sB1Formula
    .Cells(1, 2).Formula = "=Periods"
    .Cells(2, 2).Formula = "=Samples"
    'Set the starting point
    Set rTableStart = .Cells(4, 1)

    'Build the table
    rTableStart.Value = "Period"
    For r = 1 To iNper
        .Cells(rTableStart.Row + r, 1) = r
    Next r  '= 2 To iNPer + 1
    Set rPeriods = .Range(rTableStart.Offset(1, 0), rTableStart.Offset(iNper, 0))

    .Range(rTableStart.Offset(1, 0), rTableStart.Offset(r - 1, 0)).Font.Bold = True
    rTableStart.Offset(0, 1).Value = "N Samples"
    .Range(rTableStart.Offset(0, 0), rTableStart.Offset(0, 1)).Font.Bold = True

    'Set the formula to show the total number of samples
    Set rNSamplesPerTimePeriod = .Range(rTableStart.Offset(1, 1), rTableStart.Offset(iNper, 1))
    With rNSamplesPerTimePeriod
        .FormulaR1C1 = "=SUM(RC[1]:RC[" & lMaxCominbations & "])"
        .Font.Italic = True
    End With    ' .Range(rTableStart.Offset(1, 1), rTableStart.Offset(r - 1, 1))

    'Make sure that calculation is on
    If Not Application.Calculation = xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Calculation has been set to automatic to ensure proper functioning of this macro." _
            , vbInformation + vbOKOnly, "Automatic Calculation"
    End If

    bToggleMaxMin = True 'So that we start with Max
    'Set these values to the # of samples and 1, so that sampling increments inwards over successive rounds
    iMax = iNper
    iMin = 1

    'X This is what we use if the number of samples is less than then number of periods
    'As soon as samples are equal to, or more than, the number of periods, we need a new approach
    'Now do each sample
    Do While c <= lMaxCominbations _
            And Application.WorksheetFunction.Min(rNSamplesPerTimePeriod) < iNSam _
            And iMax >= iMin
        c = c + 1

        'Debug.Print "Max: " & iMax & " Min: " & iMin
        'Set to the middle category when the number of samples is even and the min and max are two apart
'        If bEvenSamples And iMax - iMin = 2 Then
'            iMax = iMax - 1
'            iMin = iMin + 1
'        End If
        'If bEvenSamples And Application.WorksheetFunction.Min(rNSamplesPerTimePeriod) = iNSam - 1 Then
        If Application.WorksheetFunction.Min(rNSamplesPerTimePeriod) = iNSam - 1 Then
            iMin = iNper / 2
            'iMax = iNPer \ 2
        End If

        If bToggleMaxMin Then           'Use Max
            'Sample from the start
            Call LabelTheColumn(rTableStart, c)
            .Range(rTableStart.Offset(1, c + 1), rTableStart.Offset(iMax, c + 1)).Value = 1
            If c > 1 Then
                c = c + 1

                'Sample to the end
                Call LabelTheColumn(rTableStart, c)
                .Range(rTableStart.Offset(iNper - iMax + 1, c + 1), rTableStart.Offset(iNper, c + 1)).Value = 1
            End If

            iMax = iMax - 1
        ElseIf Not bToggleMaxMin Then   'Use Min
            'Sample from the start
            Call LabelTheColumn(rTableStart, c)
            .Range(rTableStart.Offset(1, c + 1), rTableStart.Offset(iMin, c + 1)).Value = 1

            'Sample to the end
            c = c + 1
            Call LabelTheColumn(rTableStart, c)
            .Range(rTableStart.Offset(iNper - iMin + 1, c + 1), rTableStart.Offset(iNper, c + 1)).Value = 1

            iMin = iMin + 1
        End If  'bToggleMaxMin Then
        'Switch the toggle
        bToggleMaxMin = Not bToggleMaxMin

'        rTableStart.Offset(iNPer + 1, c + 1).FormulaArray = "=INDEX(" & rPeriods.Address(True, True, xlR1C1) & ",MIN(ROW(R[-" & iNPer & "]C:R[-1]C)*IF(ISBLANK(R[-" & iNPer & "]C:R[-1]C),MAX(" & rTableStart.Address(True, True, xlR1C1) & ")+1,1))-MIN(ROW(" & rPeriods.Address(True, True, xlR1C1) & ")))"
'        rTableStart.Offset(iNPer + 2, c + 1).FormulaArray = "=INDEX(" & rPeriods.Address(True, True, xlR1C1) & ",MAX(ROW(R[-" & (iNPer + 1) & "]C:R[-2]C)*NOT(ISBLANK(R[-" & (iNPer + 1) & "]C:R[-2]C)))-MIN(ROW(" & rPeriods.Address(True, True, xlR1C1) & "))+1)"
    Loop    'While c <= iNPer * iNSam And Application.WorksheetFunction.Min(rNSamplesPerTimePeriod.Address) < iNSam

    'I am doing this as a separate loop which is only invoked when there are more samples requested than periods
    If iNSam > iNper Then


    End If

    'Tidy up

    Set rStart = rTableStart.Offset(iNper + 1, 1)
    rStart.Value = "Start"
    Set rEnd = rStart.Offset(1, 0)
    rEnd.Value = "End"
    Set rStart = .Range(rStart.Offset(0, 1), rStart.Offset(0, 1))
    Set rEnd = .Range(rEnd.Offset(0, 1), rEnd.Offset(0, 1))

    Set rTableStart = .Range(rTableStart.Offset(1, 0), rTableStart.Offset(iNper, 0))

    '{=INDEX($A$5:$A$16,MIN(ROW(C5:C16)*IF(ISBLANK(C5:C16),MAX($A$5:$A$16)+1,1))-MIN(ROW($A$5:$A$16)))}
    rStart.FormulaArray = "=INDEX(" & rTableStart.Address(True, True, xlR1C1) & ",MIN(ROW(R[-" & iNper & "]C:R[-1]C)*IF(ISBLANK(R[-" & iNper & "]C:R[-1]C),MAX(" & rTableStart.Address(True, True, xlR1C1) & ")+1,1))-(MIN(ROW(" & rTableStart.Address(True, True, xlR1C1) & "))-1))"
'    rStart.FormulaArray = "=INDEX(" & rTableStart.Address(True, True, xlR1C1) & ",MIN(ROW(R[-" & iNPer & "]C:R[-1]C[" & c & "])*IF(ISBLANK(R[-" & iNPer & "]C:R[-1]C[" & c & "]),MAX(" & rTableStart.Address(True, True, xlR1C1) & ")+1,1))-MIN(ROW(" & rTableStart.Address(True, True, xlR1C1) & ")))"
    '{=INDEX($A$5:$A$16,MAX(ROW(C5:C16)*NOT(ISBLANK(C5:C16)))-MIN(ROW($A$5:$A$16))+1)}
    rEnd.FormulaArray = "=INDEX(" & rTableStart.Address(True, True, xlR1C1) & ",MAX(ROW(R[-" & (iNper + 1) & "]C:R[-2]C)*NOT(ISBLANK(R[-" & (iNper + 1) & "]C:R[-2]C)))-MIN(ROW(" & rTableStart.Address(True, True, xlR1C1) & "))+1)"
'    rEnd.FormulaArray = "=INDEX(" & rTableStart.Address(True, True, xlR1C1) & ",MAX(ROW(R[-" & iNPer - 1 & "]C:R[-2]C[" & c & "])*NOT(ISBLANK(R[-" & iNPer - 1 & "]C:R[-2]C[" & c & "])))-MIN(ROW(" & rTableStart.Address(True, True, xlR1C1) & "))+1)"

    .Range(rStart, rEnd).Copy .Range(rStart.Offset(0, 1), rEnd.Offset(0, c - 1))
    With .Range(rStart.Offset(0, -1), rEnd.Offset(0, c - 1)).Font
        .Bold = True
        .Italic = True
    End With
    '.Range(rTableStart, rTableStart.Offset(iNPer, c + 1)).Columns.AutoFit
    .Cells.Columns.AutoFit
    .Cells(3, 5).Select
    ActiveWindow.FreezePanes = True
End With    'ws
End Sub
Sub LabelTheColumn(ByRef rn As Range, ByRef j As Long)
    'Label the column
    With rn.Offset(0, j + 1)
        .Value = j
        With .Font
            .Italic = True
            .Bold = True
        End With
    End With

End Sub

Sub PlotAllCombinations()
    Dim per As Long
'    Dim sam As Long
    Dim r As Long
    Dim c As Long
    Dim ss As Long
    Dim pos As Long
    per = ActiveSheet.Cells(1, 2).Value
'    sam = ActiveSheet.Cells(2, 2).value
    r = 5
    c = 3
    For ss = 1 To per
        For pos = 1 To per - ss + 1
            Range(Cells(r + pos - 1, c), Cells(r + pos - 1 + ss - 1, c)).Value = 1
            c = c + 1
        Next pos
    Next ss
End Sub

