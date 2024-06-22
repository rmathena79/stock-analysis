Attribute VB_Name = "Module1"
' Roy Mathena
' Module 2 Challenge script
' Note this code needs to be in a MODULE, not just associated with a workbook or worksheet

' Column numbers, declared as named constants for readability
' Input columns (the data provided in the original spreadsheet)
Const IN_COL_TICKER As Integer = 1
Const IN_COL_OPEN As Integer = 3
Const IN_COL_CLOSE As Integer = 6
Const IN_COL_VOLUME As Integer = 7

' Overall output colomns
Const OUT_COL_FIRST As Integer = 9
Const OUT_COL_LAST As Integer = 17

' Output columns for the main results for each ticker symbol
Const OUT_COL_TICKER As Integer = 9
Const OUT_COL_QTR_CHANGE_VAL As Integer = 10
Const OUT_COL_QTR_CHANGE_PCT As Integer = 11
Const OUT_COL_VOLUME As Integer = 12

' Output columns for the little "greatest XYZ" table
Const OUT_COL_SUMMARY_LABELS As Integer = 15
Const OUT_COL_SUMMARY_TICKER As Integer = 16
Const OUT_COL_SUMMARY_VALUE As Integer = 17

' Results calculated for each ticker symbol
Type TickerResults
    Name As String
    FirstOpen As Currency ' I'm not sure this is the ideal data type, but it's sufficient
    LastClose As Currency
    ChangeValue As Currency
    ChangePercent As Double
    TotalVolume As LongLong
    LastRow As Long
End Type

' Simple helper to get the last populate cell in a column, used in a couple places below
Function GetLastRow(Col As Integer) As Long
    GetLastRow = Cells(Rows.Count, Col).End(xlUp).Row
End Function

' Erase any output cells, just for convenience
Sub Cleanup()
    Dim LastRow As Long
    Dim Col As Integer
    
    For Col = OUT_COL_FIRST To OUT_COL_LAST
        LastRow = GetLastRow(Col)
        For Row = 1 To LastRow
            Cells(Row, Col).Value = ""
            Cells(Row, Col).FormatConditions.Delete
        Next Row
    Next Col
End Sub

' Gather statistics for one ticker symbol
Function GetTickerResults(Name As String, FirstRow As Long) As TickerResults
    Const NOT_SET As Integer = -1
        
    Dim Result As TickerResults
    Dim Done As Boolean
    Dim Row As Long
    
    ' Put results to sane initial values
    Result.Name = Name
    Result.FirstOpen = NOT_SET
    Result.LastClose = NOT_SET
    Result.ChangeValue = NOT_SET
    Result.ChangePercent = NOT_SET
    Result.TotalVolume = 0
    Result.LastRow = Row
    
    ' Iterate through rows until you run out of matching names
    Done = False
    Row = FirstRow
    While Not Done
        ThisRowName = Cells(Row, IN_COL_TICKER).Value
        If ThisRowName = Name Then
            ' Check whether to save the opening price for the quarter
            If Result.FirstOpen = NOT_SET Then
                ' This is the first row for this ticker symbol,
                ' so it's the opener for the quarter
                Result.FirstOpen = Cells(Row, IN_COL_OPEN).Value
            End If
            
            ' Save the closing price every time,
            ' since we'll ultimately want the last one
            Result.LastClose = Cells(Row, IN_COL_CLOSE).Value
            
            ' Add the volume for a running total
            Result.TotalVolume = Result.TotalVolume + Cells(Row, IN_COL_VOLUME).Value
            
            ' Just keep overwriting the last row,
            ' just like with closing price
            Result.LastRow = Row
            
            ' Go to next row
            Row = Row + 1
        Else
            ' We found a different name, or blank cell, so we're done with this ticker name
            Done = True
        End If
    Wend
    
    ' Calculate overall changes
    Result.ChangeValue = Result.LastClose - Result.FirstOpen
    Result.ChangePercent = Result.ChangeValue / Result.FirstOpen
    
    GetTickerResults = Result
End Function

Sub PopulateStockSummary()
    ' Start by cleaning up from any previous runs
    Cleanup

    ' Row numbers, working down the sheet
    Dim InRow As Long
    Dim OutRow As Long
    InRow = 2
    OutRow = 1
    
    ' "Greatest so far" results, name is left blank to be populated in first pass through loop
    Dim MaxIncrease As TickerResults
    Dim MaxDecrease As TickerResults
    Dim MaxVolume As TickerResults
        
    ' Write column headers
    Cells(OutRow, OUT_COL_TICKER).Value = "Ticker"
    Cells(OutRow, OUT_COL_QTR_CHANGE_VAL).Value = "Quarterly Change"
    Cells(OutRow, OUT_COL_QTR_CHANGE_PCT).Value = "Percent Change"
    Cells(OutRow, OUT_COL_VOLUME).Value = "Total Stock Volume"
    OutRow = OutRow + 1
    
    ' Gather stats for each ticker until you hit a blank name
    Dim Name As String
    Name = Cells(InRow, IN_COL_TICKER)
    While Name <> ""
        ' Get the stats for this particular ticker symbol
        Dim Results As TickerResults
        Results = GetTickerResults(Name, InRow)
        
        ' Update greatest-so-fars
        ' Note the direction on the increase/decrease could end up wrong if there aren't any stocks
        ' that went the right way. So if all stocks decrease, the "greatest increase" will be a decrease.
        ' That seems like a reasonable way to handle that unlikely case.
        If (MaxIncrease.Name = "") Or (Results.ChangePercent > MaxIncrease.ChangePercent) Then
            MaxIncrease = Results
        End If
        If (MaxDecrease.Name = "") Or (Results.ChangePercent < MaxDecrease.ChangePercent) Then
            MaxDecrease = Results
        End If
        If (MaxVolume.Name = "") Or (Results.TotalVolume > MaxVolume.TotalVolume) Then
            MaxVolume = Results
        End If
        
        ' Write the ticker-specific results
        Cells(OutRow, OUT_COL_TICKER).Value = Name
        Cells(OutRow, OUT_COL_QTR_CHANGE_VAL).Value = Results.ChangeValue
        Cells(OutRow, OUT_COL_QTR_CHANGE_PCT).Value = Results.ChangePercent
        Cells(OutRow, OUT_COL_VOLUME).Value = Results.TotalVolume
                 
        ' Do numeric formats for those results
        Cells(OutRow, OUT_COL_QTR_CHANGE_VAL).NumberFormat = "0.00"
        Cells(OutRow, OUT_COL_QTR_CHANGE_PCT).NumberFormat = "0.00%"
        Cells(OutRow, OUT_COL_VOLUME).NumberFormat = "0"
        
        ' Do the conditional color-coding for change direction
        ' (Based on code generated by Copilot)
        With Cells(OutRow, OUT_COL_QTR_CHANGE_VAL).FormatConditions
            ' Remove any previous formatting
            .Delete
        
            ' Red (values less than 0)
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .Item(1).Interior.Color = RGB(255, 0, 0)
            
            ' White (values exactly 0)
            .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
            .Item(2).Interior.Color = RGB(255, 255, 255)
            
            ' Green (values greater than 0)
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .Item(3).Interior.Color = RGB(0, 255, 0)
        End With
        
        ' Move to next ticker symbol, if any
        InRow = Results.LastRow + 1
        Name = Cells(InRow, IN_COL_TICKER)
        OutRow = OutRow + 1
    Wend
    
    ' Write the little summary table -- I'm not bothering with a row number variable since it's a fixed size
    Cells(1, OUT_COL_SUMMARY_TICKER).Value = "Ticker"
    Cells(1, OUT_COL_SUMMARY_VALUE).Value = "Value"
    
    Cells(2, OUT_COL_SUMMARY_LABELS).Value = "Greatest % Increase"
    Cells(2, OUT_COL_SUMMARY_TICKER).Value = MaxIncrease.Name
    Cells(2, OUT_COL_SUMMARY_VALUE).Value = MaxIncrease.ChangePercent
    Cells(2, OUT_COL_SUMMARY_VALUE).NumberFormat = "0.00%"
    
    Cells(3, OUT_COL_SUMMARY_LABELS).Value = "Greatest % Decrease"
    Cells(3, OUT_COL_SUMMARY_TICKER).Value = MaxDecrease.Name
    Cells(3, OUT_COL_SUMMARY_VALUE).Value = MaxDecrease.ChangePercent
    Cells(3, OUT_COL_SUMMARY_VALUE).NumberFormat = "0.00%"
    
    Cells(4, OUT_COL_SUMMARY_LABELS).Value = "Greatest Total Volume"
    Cells(4, OUT_COL_SUMMARY_TICKER).Value = MaxVolume.Name
    Cells(4, OUT_COL_SUMMARY_VALUE).Value = MaxVolume.TotalVolume
    Cells(4, OUT_COL_SUMMARY_VALUE).NumberFormat = "0" ' The instructions show this in scientific notation, but this is better.

    ' AutoFit the output columns so everything is readable
    Dim LastRow As Long
    Dim FitRange As String
    Dim Col As Integer
    For Col = OUT_COL_FIRST To OUT_COL_LAST
        LastRow = GetLastRow(Col)
        If LastRow > 1 Then ' Don't bother if we didn't put anythin in it
            FitRange = Cells(1, Col).Address + ":" + Cells(LastRow, Col).Address
            Range(FitRange).Columns.AutoFit
        End If
    Next Col
End Sub

Sub PopulateAllSheets()
    ' Declare Current as a worksheet object variable.
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        PopulateStockSummary
    Next
End Sub

Sub CleanupAllSheets()
    ' Declare Current as a worksheet object variable.
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        Cleanup
    Next
End Sub

