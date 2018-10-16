

Option Explicit

Sub stockStats()

Dim r As Long
Dim c As Long
Dim ws As Worksheet
Dim lastRow As Long
Dim tickerName As String
Dim currentTicker As String
Dim nextTicker As String
Dim totalVol As Double
Dim summaryTableRow As Integer
Dim tickerOpen As Double
Dim tickerClose As Double
Dim yearlyChange As Double
Dim PercentChange As Double

'initialize variables
totalVol = 0




For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    summaryTableRow = 2
    
    
    tickerOpen = ws.Cells(2, 3).Value
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'loop through all the ticker
    For r = 2 To lastRow
        
        currentTicker = ws.Cells(r, 1).Value
        nextTicker = ws.Cells(r + 1, 1).Value
        
        'add to the total volume per ticker
        totalVol = totalVol + ws.Cells(r, 7).Value
        
        If (nextTicker <> currentTicker) Then
            
            'set the ticker name
            tickerName = ws.Cells(r, 1).Value
                           
            'set the ticker close
            tickerClose = ws.Cells(r, 5).Value
            
            'compute yearly ticker price change
            yearlyChange = tickerClose - tickerOpen
            
            'print ticker yearly change
            ws.Range("J" & summaryTableRow).Value = yearlyChange
            
                If (yearlyChange > 0) Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                End If

            
            'compute percent change
            If (tickerOpen = 0) Then
                PercentChange = 0
            Else
                PercentChange = yearlyChange / tickerOpen
            End If
            
            'print percent change
            ws.Range("K" & summaryTableRow).Value = PercentChange
            
            'print ticker name in summary table
            ws.Range("I" & summaryTableRow).Value = tickerName
            
            'print the totalVol to summary table
            ws.Range("L" & summaryTableRow).Value = totalVol
            
            'add 1 to the summary table row
            summaryTableRow = summaryTableRow + 1
            
            'reset the total volume
            totalVol = 0
            
            'reset the ticker close
            tickerClose = 0
            
            'grab next ticker open
            tickerOpen = ws.Cells(r + 1, 3).Value
            
        End If
        
    Next r

Next ws
    Call locateStock

End Sub

Sub locateStock()
    Dim ws As Worksheet
    Dim incPercentChange As Double
    Dim decPercentChange As Double
    Dim tickerVol As Double
    
    Dim PercentChange As Double
    Dim r As Long
    Dim lastRow As Double
    Dim stockVol As Double
    Dim highStockVol As Double
    
    'loop the worksheets
    For Each ws In Worksheets
        
        lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        incPercentChange = 0
        decPercentChange = 0
        
        highStockVol = 0
        
        For r = 2 To lastRow
            
            PercentChange = ws.Cells(r, 11).Value
            stockVol = ws.Cells(r, 12).Value
            
            If (PercentChange > incPercentChange) Then
                incPercentChange = PercentChange
                ws.Cells(2, 16).Value = incPercentChange
                ws.Cells(2, 15).Value = ws.Cells(r, 9).Value
                ws.Cells(2, 16).NumberFormat = "0%"
            End If
            
            If (PercentChange < decPercentChange) Then
                decPercentChange = PercentChange
                ws.Cells(3, 16).Value = decPercentChange
                ws.Cells(3, 15).Value = ws.Cells(r, 9).Value
                ws.Cells(3, 16).NumberFormat = "0%"
            End If
            
            If (stockVol > highStockVol) Then
                highStockVol = stockVol
                ws.Cells(4, 16).Value = highStockVol
                ws.Cells(4, 15).Value = ws.Cells(r, 9).Value
                ws.Cells(4, 16).NumberFormat = "#,##0"
            End If
    
        Next r
            
        ws.Columns("A:P").AutoFit
            
    Next ws

End Sub





