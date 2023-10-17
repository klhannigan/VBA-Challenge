Attribute VB_Name = "Module1"
Sub tickerloop()

    For Each ws In Worksheets

        Dim TickerName As String
    
        Dim TickerVolume As Double
        Dim TickerRowSummary As Integer
        Dim OpenPrice As Double
        
        TickerVolume = 0
        TickerRowSummary = 2
        OpenPrice = ws.Cells(2, 3).Value
        
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim YearlyPercentChange As Double

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              TickerName = ws.Cells(i, 1).Value

              TickerVolume = TickerVolume + ws.Cells(i, 7).Value

              ws.Range("I" & TickerRowSummary).Value = TickerName
              ws.Range("L" & TickerRowSummary).Value = TickerVolume

              ClosePrice = ws.Cells(i, 6).Value

               YearlyChange = (ClosePrice - OpenPrice)
              
              ws.Range("J" & TickerRowSummary).Value = YearlyChange

                If OpenPrice = 0 Then
                    percent_change = 0
                
                Else
                    YearlyPercentChange = YearlyChange / OpenPrice
                
                End If

              ws.Range("K" & TickerRowSummary).Value = YearlyPercentChange
              ws.Range("K" & TickerRowSummary).NumberFormat = "0.00%"
   
              TickerRowSummary = TickerRowSummary + 1

              TickerVolume = 0

              OpenPrice = ws.Cells(i + 1, 3)
            
            Else
              
              TickerVolume = TickerVolume + ws.Cells(i, 7).Value

            End If
        
        Next i

    LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To LastRowSummary
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
      
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To LastRowSummary
        
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowSummary)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowSummary)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowSummary)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub
