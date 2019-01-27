Attribute VB_Name = "Module1"
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

'You will also need to display the ticker symbol to coincide with the total volume.


Sub stock_data()

For Each ws In Worksheets


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


  Dim Ticker As String
  Dim Volume_Total As Double
  
    Volume_Total = 0


  Dim Summary_values As Integer
  Summary_values = 2
  
  ws.Cells(1, 9).Value = "Ticker"

  ws.Cells(1, 10).Value = "Total Volume"
  
  
  For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Ticker = ws.Cells(i, 1).Value

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

      ws.Range("I" & Summary_values).Value = Ticker

      ws.Range("J" & Summary_values).Value = Volume_Total

      Summary_values = Summary_values + 1
      
      Volume_Total = 0
      

   
    Else

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub
