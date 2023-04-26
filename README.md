# VBA-challenge


Through-out this assignment I received assistance from learning assistants and my course instructor. They helped me find errors in my code. I am making this acknowledgement as per the assignment guidelines and hope this suffices and meets the expected requirement you are looking for. I am pasting my full code in here as well just in case the file does not work.

Sub Headers():
Dim ws As Worksheet
  For Each ws In Worksheets
    'insert new headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

  Next ws
End Sub






Sub StockPrices():
Headers
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Volume As Double
    Dim PercentageChange As Double
    Dim Year As Integer
    Dim SummaryTable As Double
    SummaryTable = 2
    YearlyChange = 0
    Dim ws As Worksheet
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For Each ws In Worksheets
     NewTicker = 2
     For i = 2 To lastRow
    
    
        'Check is ticker is the same
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
             'Set ticker name and Volume
             Ticker = ws.Cells(i, 1).Value
             Volume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(NewTicker, 7), ws.Cells(i, 7)))
            
             'Open and Close Ticker price
             OpenPrice = ws.Cells(NewTicker, 3).Value
             ClosePrice = ws.Cells(i, 6).Value
             NewTicker = (i + 1)
             
             'Find Yearly and percent Changes
             YearlyChange = ClosePrice - OpenPrice
             PercentageChange = YearlyChange / OpenPrice
             ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
             
             'Where all the data goes
             ws.Range("I" & SummaryTable).Value = Ticker
             ws.Range("J" & SummaryTable).Value = YearlyChange
             ws.Range("K" & SummaryTable).Value = PercentageChange
             ws.Range("L" & SummaryTable).Value = Volume
             If ws.Range("J" & SummaryTable).Value > 0 Then
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
             ElseIf ws.Range("J" & SummaryTable).Value < 0 Then
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
             End If
             If ws.Range("K" & SummaryTable).Value > 0 Then
                ws.Range("K" & SummaryTable).Interior.ColorIndex = 4
             ElseIf ws.Range("J" & SummaryTable).Value < 0 Then
                ws.Range("K" & SummaryTable).Interior.ColorIndex = 3
             End If
             SummaryTable = SummaryTable + 1
            
            
            End If
        
    
        Next i
    
    SummaryTable = 2
    
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    ws.Range("Q4").NumberFormat = "0"
    gt_increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P2") = Cells(gt_increase_index + 1, 9)
    gt_decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P3") = Cells(gt_decrease_index + 1, 9)
    gt_volume_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    ws.Range("P4") = Cells(gt_volume_index + 1, 9)
    
    Next ws
    
End Sub



