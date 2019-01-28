Attribute VB_Name = "YearlyStockCalculator"
' Purpose : For each sheet within the workbook, calculates the yearly results of the stocks that are on the sheet and updates
'               the results at the top of that sheet.
'
'               The code has the following assumptions:
'               A The format of the sheets are the same throughout the workbook; the data is contained within the same columns
'               B The start and end date can change based on the year; this is due to holidays/weekends when there are no stocks traded
'               C When the start date had a opening value of zero, percent changed is not calculated and set to zero.
'               D The rows are grouped by a ticket and sorted by date with the first date at the top
'
' Developer : Scott McEachern
'
' Date : January 25, 2019
'

'- Consts

' Index of the row to place the result header information
Private Const mc_headerRowIndex As Integer = 1

' Index of the column that contains the ticker text
Private Const mc_sourceTickerColumnIndex As Integer = 1

' Index of the column that contains the date that the row contains information for
Private Const mc_sourceDateColumnIndex As Integer = 2

' Index of the column that contains the opening stock price for the day
Private Const mc_sourceOpenPriceColumnIndex As Integer = 3

' Index of the column that contains the closing stock price for the day
Private Const mc_sourceClosePriceColumnInex As Integer = 6

' Index of the column that contains the volumn of the stock for the day
Private Const mc_sourceVolumeColumnIndex As Integer = 7

' Index of the column that is to be populated with the ticker name to display the results
Private Const mc_resultTickerColumnIndex As Integer = 9

' Index of the column that is populated with the yearly change results
Private Const mc_resultYearlyChangeColumnIndex As Integer = 10

' Index of the column that contains the percentage changed results
Private Const mc_resultPercentageChangedColumnIndex As Integer = 11

' Index of the column that contains the total stock volumn for the ticker
Private Const mc_resultTotalStockVolumnColumnIndex As Integer = 12

' Index of the column that contains the label for the greatest results
Private Const mc_greatestLabelColumnIndex As Integer = 15

' Index of the column that is to contain the ticker for the years greatest results
Private Const mc_greatestTickerColumnIndex As Integer = 16

' Index of the column that contains the value for the greats results
Private Const mc_greatestValueColumnIndex As Integer = 17

' Format to display the percent change column
Private Const mc_resultPercentageChangedFormat As String = "0.00%"

' Format to display the yearly change column
Private Const mc_resultYearlyChangeFormat As String = "#0.00000000"




Public Sub Start()
' Purpose : Starting point for the yearly calculation; loop through all of the worksheets within the Excel document

On Error GoTo PROC_ERR:

    Dim numWorksheets As Integer
    Dim worksheetCounter As Integer
    Dim startTime As Date

    '- Start Timer
    startTime = Time


    '- Determine Number of Sheets
    numWorksheets = ThisWorkbook.Worksheets.Count
    
    
    '- Process each worksheet
    For worksheetCounter = 1 To numWorksheets
        CalculateForWorksheet (worksheetCounter)
    Next worksheetCounter


    '- Complete Timer
    Dim totalSeconds As String
    totalSeconds = DateDiff("s", startTime, Time)
    
    
    '- Inform User
    MsgBox ("Completed calculations; total processing time: " + totalSeconds + " seconds")
    
    
    Exit Sub

PROC_ERR:
    MsgBox (Err.Description)
End Sub


Private Sub CalculateForWorksheet(worksheetIndex As Integer)
' Purpose : For the worksheet provided, calculate the summary information for each ticker and deterine the greatest tickers
'
' Accepts : worksheetIndex      Index of the worksheet that is to have the calculations performed

    Dim numTotalRows As Double
    Dim rowCounter As Double
    Dim sourceTicker As TickerInfo
    Dim isFirstRow As Boolean
    Dim previousClosePrice As Double
    Dim greatestTicker As GreatestTickerInfo
    
    
    '- Prepare For Results
    PrepareWorksheetForResults (worksheetIndex)


    '- Determine Number of Rows
    numTotalRows = GetNumberOfRowsForColumn(worksheetIndex, mc_sourceTickerColumnIndex)
    
    
    '- Loop Through Rows
    isFirstRow = True
    previousClosePrice = 0
    Set greatestTicker = New GreatestTickerInfo
    
    For rowCounter = 2 To numTotalRows
    
        If (isFirstRow = True) Then
        
            '- First Row, Prepare new sourceTicker
            isFirstRow = False
            
            Set sourceTicker = StartNewTickerInfo(worksheetIndex, rowCounter)
        
        ElseIf (sourceTicker.Ticker <> Worksheets(worksheetIndex).Cells(rowCounter, mc_sourceTickerColumnIndex).value) Then
            
            '- Complete Ticker
            '   The last row contains the closing price for the ticket
            sourceTicker.CompleteTicker (previousClosePrice)
            
            
            '- Check for Greatest Ticker
            Call CheckForGreatestTicker(sourceTicker, greatestTicker)
            
            
            '- Update Results
            Call UpdateResultsForTicker(worksheetIndex, sourceTicker)
            
            
            '- Prepare new sourceTicker
            Set sourceTicker = StartNewTickerInfo(worksheetIndex, rowCounter)
        End If
        
        
        '- Update Volume
        sourceTicker.AddVolume (Worksheets(worksheetIndex).Cells(rowCounter, mc_sourceVolumeColumnIndex).value)
    
    
        '- Set Previous Close Price
        '   Used when reach end of rows for ticker
        previousClosePrice = Worksheets(worksheetIndex).Cells(rowCounter, mc_sourceClosePriceColumnInex).value
    
    Next rowCounter
    
    
    '- Last Row
    sourceTicker.CompleteTicker (previousClosePrice)
        
    Call UpdateResultsForTicker(worksheetIndex, sourceTicker)

    
    '- Update Greatest Ticker
    Call UpdateGreatestTickerResults(worksheetIndex, greatestTicker)
    
End Sub

Private Sub UpdateGreatestTickerResults(worksheetIndex As Integer, greatestTicker As GreatestTickerInfo)
' Purpose : Updates the sheet with the greatest ticker information
'
' Accepts : worksheetIndex      Index of the worksheet currently in use
'                greatestTicker         contains the results for the sheet

    '- Greatest Increase
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 1, mc_greatestLabelColumnIndex).value = "Greatest % Increase"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 1, mc_greatestTickerColumnIndex).value = greatestTicker.GreatestIncreaseTicker
    
    With Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 1, mc_greatestValueColumnIndex)
        .value = greatestTicker.GreatestIncreaseValue
        .NumberFormat = mc_resultPercentageChangedFormat
    End With
    
    
    '- Greatest Decrease
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 2, mc_greatestLabelColumnIndex).value = "Greatest % Decrease"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 2, mc_greatestTickerColumnIndex).value = greatestTicker.GreatestDecreaseTicker
    
    With Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 2, mc_greatestValueColumnIndex)
        .value = greatestTicker.GreatestDecreaseValue
        .NumberFormat = mc_resultPercentageChangedFormat
    End With
    
    
    '- Greatest Volume
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 3, mc_greatestLabelColumnIndex).value = "Greatest Total Volume"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 3, mc_greatestTickerColumnIndex).value = greatestTicker.GreatestVolumeTicker
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex + 3, mc_greatestValueColumnIndex).value = greatestTicker.GreatestVolumeValue


End Sub

Private Sub CheckForGreatestTicker(sourceTicker As TickerInfo, greatestTicker As GreatestTickerInfo)
' Purpose : Checks for the greatest changes and when the source is found to have it, then updates the greatestTicker
'                that stores this information
'
' Accepts : sourceTicker        reference to the ticker metadata that is to be compared to
'                greatestTicker     contains metadata on the greatest changes for the worksheet

    '- Check: Greatest Increase
    If (sourceTicker.YearlyChangePercentage > greatestTicker.GreatestIncreaseValue) Then
        greatestTicker.GreatestIncreaseTicker = sourceTicker.Ticker
        greatestTicker.GreatestIncreaseValue = sourceTicker.YearlyChangePercentage
    End If
    
    
    '- Check: Greatest Decrease
    If (sourceTicker.YearlyChangePercentage < greatestTicker.GreatestDecreaseValue) Then
        greatestTicker.GreatestDecreaseTicker = sourceTicker.Ticker
        greatestTicker.GreatestDecreaseValue = sourceTicker.YearlyChangePercentage
    End If
    
    
    '- Check: Greatest Volume
    If (sourceTicker.TotalVolume > greatestTicker.GreatestVolumeValue) Then
        greatestTicker.GreatestVolumeTicker = sourceTicker.Ticker
        greatestTicker.GreatestVolumeValue = sourceTicker.TotalVolume
    End If
    
End Sub

Private Function StartNewTickerInfo(worksheetIndex As Integer, rowCounter As Double) As TickerInfo
' Purpose : Prepare new tickerInfo, assumption is that the row provided is the first row found for that ticker
'
' Accepts : worksheetIndex      Index of the worksheet getting results for
'                rowCounter             Index of the row that is the start of the ticker
'
' Returns : sourceTicker            newly created instance that is updated with ticker and FirstDayPrice

    Dim sourceTicker As TickerInfo
    
    Set sourceTicker = New TickerInfo
    
    
    '- Set Ticker Name
    sourceTicker.Ticker = Worksheets(worksheetIndex).Cells(rowCounter, mc_sourceTickerColumnIndex).value
    
    '- Set Open Price
    sourceTicker.FirstDayPrice = Worksheets(worksheetIndex).Cells(rowCounter, mc_sourceOpenPriceColumnIndex).value
    
    
    Set StartNewTickerInfo = sourceTicker
End Function


Private Sub UpdateResultsForTicker(worksheetIndex As Integer, sourceTicker As TickerInfo)
' Purpose : Sets the results for the ticker on the worksheet
'
' Accepts : worksheetIndex  Index of the worksheet to update with results for ticker
'                sourceTicker       contains the information collected from the worksheet for individual ticker
    
    Dim numRows As Double
    Dim resultRowIndex As Double
    Dim yearlyChangeColorIndex As Integer
    
    
    '- Determine Rows
    numRows = GetNumberOfRowsForColumn(worksheetIndex, mc_resultTickerColumnIndex)
    
    resultRowIndex = (numRows + 1)
    
    
    '- Set Results: Ticker
    Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultTickerColumnIndex).value = sourceTicker.Ticker
    
    
    '- Set Results: Yearly Change
    '   Only update the fill color of the cell when not zero
    With Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultYearlyChangeColumnIndex)
        .value = sourceTicker.YearlyChangeActual
        .NumberFormat = mc_resultYearlyChangeFormat
    End With
    
    If (sourceTicker.YearlyChangeActual <> 0) Then
        If (sourceTicker.YearlyChangeActual > 0) Then
            ' Green Color
            yearlyChangeColorIndex = 4
        Else
            ' Red Color
            yearlyChangeColorIndex = 3
        End If
        
        Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultYearlyChangeColumnIndex).Interior.ColorIndex = yearlyChangeColorIndex
    End If
    

    '- Set Result: Percentage Changed
    Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultPercentageChangedColumnIndex).value = sourceTicker.YearlyChangePercentage
    
    Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultPercentageChangedColumnIndex).NumberFormat = mc_resultPercentageChangedFormat
    
    
    '- Set Result: Yearly Volume
    Worksheets(worksheetIndex).Cells(resultRowIndex, mc_resultTotalStockVolumnColumnIndex).value = sourceTicker.TotalVolume
    
End Sub


Private Function GetNumberOfRowsForColumn(worksheetIndex As Integer, columnIndex As Integer) As Double
' Purpose : Determines the number of rows for the column provided. Method based on StackOverflow:
'                   https://stackoverflow.com/questions/21554059/how-to-get-the-row-count-in-excel-vba
'
' Accepts : worksheetIndex      Index of the worksheet to calculate number of rows
'                columnIndex          Index of the column to determine the number of rows

    Dim numTotalRows As Double
    
    numTotalRows = Worksheets(worksheetIndex).Cells(Worksheets(worksheetIndex).Rows.Count, columnIndex).End(xlUp).Row
    
    GetNumberOfRowsForColumn = numTotalRows
End Function


Private Sub PrepareWorksheetForResults(worksheetIndex As Integer)
' Purpose : Prepares the worksheet by setting the headers for the results and clearing existing results; this is necessary
'                for when building out the code and running the processing multiple times
'
' Accepts : worksheetIndex      Index of the worksheet to update

    '- Clear Results
    Worksheets(worksheetIndex).Columns(mc_resultTickerColumnIndex).Clear
    
    Worksheets(worksheetIndex).Columns(mc_resultYearlyChangeColumnIndex).Clear
    
    Worksheets(worksheetIndex).Columns(mc_resultPercentageChangedColumnIndex).Clear
    
    Worksheets(worksheetIndex).Columns(mc_resultTotalStockVolumnColumnIndex).Clear
    
    Worksheets(worksheetIndex).Columns(mc_greatestTickerColumnIndex).Clear
    
    Worksheets(worksheetIndex).Columns(mc_greatestValueColumnIndex).Clear
    

    '- Set Result Headers
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_resultTickerColumnIndex).value = "Ticker"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_resultYearlyChangeColumnIndex).value = "Yearly Change"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_resultPercentageChangedColumnIndex).value = "Percent Change"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_resultTotalStockVolumnColumnIndex).value = "Total Stock Volume"


    '- Set Result Column Width
    Worksheets(worksheetIndex).Columns(mc_resultYearlyChangeColumnIndex).ColumnWidth = 20
    
    Worksheets(worksheetIndex).Columns(mc_resultPercentageChangedColumnIndex).ColumnWidth = 20
    
    Worksheets(worksheetIndex).Columns(mc_resultTotalStockVolumnColumnIndex).ColumnWidth = 20


    '- Set Greatest Headers
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_greatestTickerColumnIndex).value = "Ticker"
    
    Worksheets(worksheetIndex).Cells(mc_headerRowIndex, mc_greatestValueColumnIndex).value = "Value"
    
    
    '- Set Greatest Column Width
    Worksheets(worksheetIndex).Columns(mc_greatestLabelColumnIndex).ColumnWidth = 20
    
    Worksheets(worksheetIndex).Columns(mc_greatestTickerColumnIndex).ColumnWidth = 10
    
    Worksheets(worksheetIndex).Columns(mc_greatestValueColumnIndex).ColumnWidth = 20
End Sub




