Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks and take the following info

'Yearly change from what the stock opened the year at to what the closing price was.

'The percent change from the what it opened the year at to what it closed.

'The total Volume of the stock

'Ticker Symbol


Sub stock_data()

For Each ws In Worksheets


    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim Volume_Total As Double
Dim Count As Integer
  
    Volume_Total = 0
    Count = 0

Dim Summary_values As Integer
    Summary_values = 2
  
    ws.Cells(1, 9).Value = "Ticker"

    ws.Cells(1, 12).Value = "Total Volume"
  
    ws.Cells(1, 10).Value = "Yearly Change"
  
    ws.Cells(1, 11).Value = "Percent Change"
  

  
  For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value
        
        ws.Range("I" & Summary_values).Value = Ticker
      
      date_closed = ws.Cells(i, 6).Value
      
      date_opened = ws.Cells(i - Count, 3).Value
      
      Yearly_Change = date_closed - date_opened

        ws.Range("J" & Summary_values).Value = Yearly_Change

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

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

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
      Count = Count + 1

    End If

  Next i
  
Next ws

End Sub
