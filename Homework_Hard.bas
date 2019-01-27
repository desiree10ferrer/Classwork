Attribute VB_Name = "Module1"
'Your solution will include everything from the moderate challenge
'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
Sub stock_data()

For Each ws In Worksheets


    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker As String
Dim Volume_Total As Double
Dim Count As Integer
  
    Volume_Total = 0
    Count = 0

Dim Summary_values As Integer
    Summary_values = 2
  
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 15).Value = "Ticker"
    
    ws.Cells(1, 16).Value = "Value"

    ws.Cells(1, 12).Value = "Total Volume"
  
    ws.Cells(1, 10).Value = "Yearly Change"
  
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    
    ws.Cells(4, 14).Value = "Greatest Total Volume"
  
  For I = 2 To lastrow

    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ticker = ws.Cells(I, 1).Value
        
        ws.Range("I" & Summary_values).Value = ticker
      
      date_closed = ws.Cells(I, 6).Value
      
      date_opened = ws.Cells(I - Count, 3).Value
      
      Yearly_Change = date_closed - date_opened

        ws.Range("J" & Summary_values).Value = Yearly_Change

      Volume_Total = Volume_Total + ws.Cells(I, 7).Value

        ws.Range("L" & Summary_values).Value = Volume_Total
        
    If (Yearly_Change >= 0) Then
        
        ws.Range("J" & Summary_values).Interior.ColorIndex = 4
        
        Else
        ws.Range("J" & Summary_values).Interior.ColorIndex = 3
        
        End If

    If date_opened = 0 Then
        ws.Range("K" & Summary_values) = "N/A"
      
        Else
    
    Yearly_Percentage = (Yearly_Change / date_opened) * 100 & "%"
        ws.Range("K" & Summary_values).Value = Yearly_Percentage
      
      End If
      
    Summary_values = Summary_values + 1
      
    Volume_Total = 0
      
    Count = 0
      
        
    Else

      Volume_Total = Volume_Total + ws.Cells(I, 7).Value
      
      Count = Count + 1

    End If

  Next I

    max = Application.WorksheetFunction.max(ws.Columns("K")) * 100 & "%"

        ws.Cells(2, 16).Value = max
        ws.Range("O2") = "=index(I:I,match(P2,K:K,0))"

    MIN = Application.WorksheetFunction.MIN(ws.Columns("K")) * 100 & "%"
        ws.Cells(3, 16).Value = MIN
        ws.Range("O3") = "=index(I:I,match(P3,K:K,0))"
        
    maxl = Application.WorksheetFunction.max(ws.Columns("L"))
        ws.Cells(4, 16).Value = maxl
        ws.Range("O4") = "=index(I:I,match(P4,L:L,0))"
Next ws

End Sub
