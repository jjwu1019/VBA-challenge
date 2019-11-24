Attribute VB_Name = "Module1"
Sub vbahomework():


Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    

Dim ticker As String
Dim volumn As Double
Dim yearly_change_open As Double
Dim yearly_change_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim lastrow As Double
Dim i As Double

Dim summary_table_row As Integer
summary_table_row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Volumn"

volumn = 0
yearly_change = 0
percent_change = 0
yearly_change_close = 0
yearly_change_open = 0

For i = 2 To lastrow
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        yearly_change_open = Cells(i, 3).Value
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        volumn = volumn + Cells(i, 7).Value
        yearly_change_close = Cells(i, 6).Value
        
        yearly_change = yearly_change_close - yearly_change_open
        If yearly_change_open <> 0 Then
            percent_change = yearly_change_close / yearly_change_open - 1
        ElseIf yearly_change_open = 0 Then
            percent_change = 0
        End If
        
        Range("I" & summary_table_row).Value = ticker
        Range("J" & summary_table_row).Value = yearly_change
        Range("K" & summary_table_row).Value = percent_change
        Range("L" & summary_table_row).Value = volumn
        
        
        
        
        summary_table_row = summary_table_row + 1
        
        volumn = 0
        yearly_change = 0
        percent_change = 0
        yearly_change_close = 0
        yearly_change_open = 0
    Else
        volumn = volumn + Cells(i, 7).Value
        
        
    End If
Next i

Columns("K").NumberFormat = "0.00%"

    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False


Range("P1") = "Ticker"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"

Dim greatest_increase_percent As Double
Dim greatest_decrease_percent As Double
Dim greatest_increase_ticker As String
Dim greatest_decrease_ticker As String
Dim greatest_volume_value As Double
Dim greatest_volume_ticker As String
Dim lastrow2 As Double
Dim j As Double
Dim k As Double

lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

greatest_increase_percent = Cells(2, 11).Value
greatest_decrease_percent = Cells(2, 11).Value
greatest_increase_ticker = Cells(2, 9).Value
greatest_decrease_ticker = Cells(2, 9).Value

For j = 3 To lastrow2
    If Cells(j, 11).Value > greatest_increase_percent Then
        greatest_increase_percent = Cells(j, 11).Value
        greatest_increase_ticker = Cells(j, 9).Value
    ElseIf Cells(j, 11).Value < greatest_decrease_percent Then
        greatest_decrease_percent = Cells(j, 11).Value
        greatest_decrease_ticker = Cells(j, 9).Value
    End If
Next j

Range("P2") = greatest_increase_ticker
Range("Q2") = greatest_increase_percent
Range("P3") = greatest_decrease_ticker
Range("Q3") = greatest_decrease_percent

Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

greatest_volume_value = Cells(2, 12).Value
greatest_volume_ticker = Cells(2, 9).Value

For k = 3 To lastrow2
    If Cells(k, 12).Value > greatest_volume_value Then
        greatest_volume_value = Cells(k, 12).Value
        greatest_volume_ticker = Cells(k, 9).Value
    End If
Next k

Range("P4") = greatest_volume_ticker
Range("Q4") = greatest_volume_value


Next

starting_ws.Activate

End Sub



