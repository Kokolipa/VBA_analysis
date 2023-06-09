    Sub stock_class()
        Dim ws As Worksheet
        Dim lastRow As Long

        ' Disable screen updating and events
        ' --------------------------------------------------
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        For Each ws In ThisWorkbook.Worksheets
            ' Counting the amount of rows within the sheet
            ' --------------------------------------------------
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Cells(1, 8).Value = "Ticker"
            ws.Cells(1, 9).Value = "Yearly Change"
            ws.Cells(1, 10).Value = "Percent Change"
            ws.Cells(1, 11).Value = "Total Stock Volume"

            Dim second_row As Long
            second_row = 2
            
            ' Define the ranges
            ' -----------------------------
            Dim ticker_name As String
            Dim openingPrice As Double
            Dim closingPrice As Double
            Dim yearly_change As Double
            Dim percent_change As Double
            Dim total_stock_volume As LongLong

            total_stock_volume = 0
            Dim isFirstTicker As Boolean
            isFirstTicker = True

            For i = 2 To lastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker_name = ws.Cells(i, 1).Value
                    ws.Range("H" & second_row).Value = ticker_name

                    closingPrice = ws.Cells(i, 6).Value
                    yearly_change = closingPrice - openingPrice
                    ws.Range("I" & second_row).Value = yearly_change
                    
                    ' Apply Conditional Formatting
                    ' -----------------------------
                    If yearly_change > 0 Then
                        ws.Range("I" & second_row).Interior.Color = RGB(0, 255, 0)
                    ElseIf yearly_change = 0 Then
                        ws.Range("I" & second_row).Interior.Color = RGB(247, 203, 57)
                    Else
                        ws.Range("I" & second_row).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    ' Calculating the percent change column
                    ' -----------------------------
                    percent_change = (yearly_change / openingPrice) * 100
                    ws.Range("J" & second_row).Value = percent_change
                    
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    ws.Range("K" & second_row).Value = total_stock_volume
                    ws.Range("K:K").Columns.AutoFit
                    
                    second_row = second_row + 1
                    
                    ' Resetting the total_stock_volume value
                    ' -----------------------------
                    total_stock_volume = 0
                    isFirstTicker = True
                Else
                    ' Returning the first instance of the openingPrice value
                    ' -----------------------------
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    If isFirstTicker Then
                        openingPrice = ws.Cells(i, 3).Value
                        isFirstTicker = False
                    End If
                    
                End If

                    
            Next i

            ' Setting the bonus table
            ' --------------------------
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"

            Dim increase As Double
            Dim decrease As Double
            Dim total_volume As Longlong
            
            increase = WorksheetFunction.Max(ws.Range("J:J"))
            ws.Range("P2").value = increase
            decrease = WorksheetFunction.Min(ws.Range("J:J"))
            ws.Range("P3").value = decrease
            total_volume = WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("P4").value = total_volume
            ws.Range("P:P").Columns.AutoFit

            ' Returning the respective ticker_name for each of the values in the bonus table
            ' --------------------------------------------------
            For i = 2 To lastRow
                If ws.Cells(i, 10).Value = increase Then
                    ws.Range("O2").Value = ws.Cells(i, 8).Value
                ElseIf ws.Cells(i, 10).Value = decrease Then
                    ws.Range("O3").Value = ws.Cells(i, 8).Value
                ElseIf ws.Cells(i, 11).Value = total_volume Then
                    ws.Range("O4").Value = ws.Cells(i, 8).Value
                End If
            Next i


            
            ' Re-enable screen updating and events
            ' --------------------------------------------------
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
        Next ws
    End Sub






