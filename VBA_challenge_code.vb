Sub StocktestA()
For Each ws In Worksheets

    'define variables
    Dim Ticker As String
    Dim Stock_Total_Volume As Double
    Dim Open_price As Double
    Dim Close_price As Double

    'set up summary table for data
    Dim summary_table_row As Integer
    summary_table_row = 2

    'add title headers to columns in worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Stock_Total_Volume"

    'determine last row of each worksheet
    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For a = 2 To lastRow
        'determine if there is an opening price, if no opening price, skip row
        If (ws.Cells(a, 3).Value = 0) Then
            'determine ticker
            If (ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value) Then
                Ticker = ws.Cells(a, 1).Value
            End If
        'determine stock total volume
        ElseIf (ws.Cells(a + 1, 1).Value = ws.Cells(a, 1).Value) Then
            Stock_Total_Volume = Stock_Total_Volume + ws.Cells(a, 7).Value
            'determine Open price
            If (ws.Cells(a - 1, 1).Value <> ws.Cells(a, 1).Value) Then
                Open_price = ws.Cells(a, 3).Value
            End If
        Else
            'set ticker
            Ticker = ws.Cells(a, 1).Value
            'add to stock total volume
            Stock_Total_Volume = Stock_Total_Volume + ws.Cells(a, 7).Value
            'set Close price
            Close_price = ws.Cells(a, 6).Value
            'add ticker and stock total volume in summary table row
            ws.Cells(summary_table_row, 9).Value = Ticker
            ws.Cells(summary_table_row, 12).Value = Stock_Total_Volume
            'to avoid dividing by zero
            If (Stock_Total_Volume > 0) Then
                'add yearly change in summary table row
                ws.Cells(summary_table_row, 10).Value = Close_price - Open_price
                    'change color to green if > 0, else red
                    If (ws.Cells(summary_table_row, 10).Value > 0) Then
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                    End If
                'add percent change in summary table row
                ws.Cells(summary_table_row, 11).Value = ws.Cells(summary_table_row, 10).Value / Open_price
            Else
                'set yearly change and percent change to zero if no stock data
                ws.Cells(summary_table_row, 10).Value = 0
                ws.Cells(summary_table_row, 11).Value = 0
            End If
            'set cell format to percent for percent change
            ws.Cells(summary_table_row, 11).Style = "percent"
            'reset stock total volume
            Stock_Total_Volume = 0
            'next summary row
            summary_table_row = summary_table_row + 1
        End If
    Next a

    'define greatest total volume
    Dim Greatest_Total_Volume As Double

    'add labels to data in worksheet
    ws.Range("N2").Value = "Greatest_Total_Volume"
    ws.Range("N3").Value = "Greatest_Percent_Increase"
    ws.Range("N4").Value = "Greatest_Percent_Decrease"
    ws.Range("O1").Value = "Ticker_Name"
    ws.Range("P1").Value = "Value"

    'set baseline for greatest total volume
    Greatest_Total_Volume = 0


    'offset summaryRow to equal number of ticker symbols
    summary_table_row = summary_table_row - 2

    'determine greatest total volume
    For a = 2 To summary_table_row
        If (ws.Cells(a, 12).Value > Greatest_Total_Volume) Then
            Greatest_Total_Volume = ws.Cells(a, 12).Value

            'add ticker in table
            ws.Range("O2").Value = ws.Cells(a, 9).Value
        End If
    Next a

    'add greatest total volume in table
    ws.Range("P2").Value = Greatest_Total_Volume

    'define variables for percent increase and percent decrease
    Dim Percent_Increase As Double
    Dim Percent_Decrease As Double

    'set baseline for greatest percent increase and greatest percent decrease
    Percent_Increase = 0
    Percent_Decrease = 0

    For a = 2 To summary_table_row
        'determine value for greatest percent increase
        If (ws.Cells(a, 11).Value > Percent_Increase) Then
            Percent_Increase = ws.Cells(a, 11).Value

            'add ticker in table for greatest percent increase
            ws.Range("O3") = ws.Cells(a, 9).Value
        'determine value for greatest percent decrease
        ElseIf (ws.Cells(a, 11).Value < Percent_Decrease) Then
            Percent_Decrease = ws.Cells(a, 11).Value

            'add ticker in table for greatest percent decrease
            ws.Range("O4").Value = ws.Cells(a, 9).Value
        End If
    Next a

    'add greatest percent increase and greastest decrease value in table
    ws.Range("P3").Value = Percent_Increase
    ws.Range("P4").Value = Percent_Decrease

    'set cell format to percent
    ws.Range("P3").Style = "percent"
    ws.Range("P4").Style = "percent"


Next ws

End Sub

