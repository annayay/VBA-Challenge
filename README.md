# VBA-Challenge
Sub Create_Column_Headers_Ticker_Total()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim TickerVar As String
    Dim c As Long
    Dim x As Long
    Dim Ticker_Total As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim QuarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestIncreasePercent As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecreasePercent As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double

    On Error Resume Next ' Enable error handling

    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        c = 2
        x = 2
        Ticker_Total = 0
        openPrice = 2

        ' Find the greatest values
        greatestIncreaseTicker = ""
        greatestIncreasePercent = 0
        greatestDecreaseTicker = ""
        greatestDecreasePercent = 0
        greatestVolumeTicker = ""
        greatestVolume = 0

        ' Loop through the rows in column A
        While ws.Cells(c, 1).Value <> ""
            TickerVar = ws.Cells(c, 1).Value

            ' Calculate total stock volume for each ticker
            Ticker_Total = Ticker_Total + ws.Cells(c, 7).Value

            ' Check if the next row has a different ticker symbol or if it's the last row
            If ws.Cells(c + 1, 1).Value <> TickerVar Or c = lastRow Then
                ' Output Ticker in column I
                ws.Cells(x, 9).Value = TickerVar

                ' Output Total Volume in column L
                ws.Cells(x, 12).Value = Ticker_Total

                ' Find the closing price based on the last row for the ticker
                closePrice = ws.Cells(c, 6).Value ' Close price for the quarter

                ' Find the opening price based on the first row for the ticker
                openPrice = ws.Cells(openPrice, 3).Value ' Open price for the quarter

                ' Calculate Quarterly Change and Percentage Change
                QuarterlyChange = closePrice - openPrice
                percentChange = IIf(openPrice <> 0, (closePrice - openPrice) / openPrice, 0)

                ' Output Quarterly Change in column J
                ws.Cells(x, 10).Value = QuarterlyChange

                ' Output Percent Change in column K
                ws.Cells(x, 11).Value = percentChange

                ' Apply conditional formatting after outputting values
                ApplyConditionalFormatting ws.Cells(x, 10), QuarterlyChange

                ' Update greatest values
                If percentChange > greatestIncreasePercent Then
                    greatestIncreaseTicker = TickerVar
                    greatestIncreasePercent = percentChange
                End If

                If percentChange < greatestDecreasePercent Then
                    greatestDecreaseTicker = TickerVar
                    greatestDecreasePercent = percentChange
                End If

                If Ticker_Total > greatestVolume Then
                    greatestVolumeTicker = TickerVar
                    greatestVolume = Ticker_Total
                End If

                ' Reset Ticker_Total for the next ticker symbol
                Ticker_Total = 0
                x = x + 1
                openPrice = c + 1
            End If

            c = c + 1
        Wend

        ' Find the stock with the greatest percentage increase
        Dim greatestIncreaseRow As Long
        Dim greatestIncreaseValue As Double

        greatestIncreaseRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Columns(11)), ws.Columns(11), 0)
        greatestIncreaseValue = ws.Cells(greatestIncreaseRow, 11).Value

        ' Populate the values in cells O2, P2, and Q2
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = ws.Cells(greatestIncreaseRow, 9)

        ' Populate the values in cells O2, P2, and Q2
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = ws.Cells(greatestIncreaseRow, 9).Value ' Ticker of Greatest % Increase stock
        ws.Cells(2, 17).Value = greatestIncreaseValue ' % Value of the Greatest % Increase
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecreasePercent

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume


        ' Apply conditional formatting to highlight the Greatest % Increase row
        ApplyConditionalFormatting ws.Cells(greatestIncreaseRow, 10), ws.Cells(greatestIncreaseRow, 10).Value

        ws.Cells(x + 1, 15).Value = "Greatest % Decrease"
        ws.Cells(x + 1, 16).Value = greatestDecreaseTicker
        ws.Cells(x + 1, 17).Value = greatestDecreasePercent

        ws.Cells(x + 2, 15).Value = "Greatest Total Volume"
        ws.Cells(x + 2, 16).Value = greatestVolumeTicker
        ws.Cells(x + 2, 17).Value = greatestVolume

        ws.Cells(x, 9).Value = "Total"
        ws.Cells(x, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, 12), ws.Cells(x - 1, 12)))

        ws.Columns(10).AutoFit
        ws.Columns(11).AutoFit
        ws.Columns(9).AutoFit
        ws.Columns(12).AutoFit

        ws.Columns(10).NumberFormat = "0.00"
        ws.Columns(11).NumberFormat = "0.00%"

        ws.Columns(16).AutoFit
        ws.Columns(17).AutoFit
        ws.Range("O:P").EntireColumn.AutoFit ' AutoFit columns O and P
        ws.Columns(17).NumberFormat = "0.00%" ' Format column 17 to display values with 2 decimal places
    Next ws

    On Error GoTo 0 ' Disable error handling
End Sub

Sub ApplyConditionalFormatting(rng As Range, QuarterlyChange As Double)
    If QuarterlyChange > 0 Then
        rng.Interior.Color = RGB(0, 255, 0) ' Green for positive change
    ElseIf QuarterlyChange < 0 Then
        rng.Interior.Color = RGB(255, 0, 0) ' Red for negative change
    End If
End Sub
