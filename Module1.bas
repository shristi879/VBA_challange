Attribute VB_Name = "Module1"
Sub stock_results()

     'create loop to add headers to each sheet
     For Each ws In Worsheets
     
     'create variables
     Dim WorsheetName As String
     WorksheetName = ws.Name
     Dim Total_volume As Double
     Total_volume = 0
     Dim Ticker_Name As String
     Ticker_Name = " "
     Dim start_Price As Double
     start_Price = 0
     Dim Percent_change As Double
     
     
     'create column headers for Required data and add to each worksheet
     ws.Cells(1, 10).Value = "Ticker"
     ws.Cells(1, 11).Value = "Yearly Change"
     ws.Cells(1, 12).Value = "Percentage Change"
     ws.Cells(1, 13).Value = "Total Stock volume"
     
     
     'create column headers for binus features
     ws.Cells(1, 17).Value = "Ticker"
     ws.Cells(1, 18).Value = "value"
     
     
    'create Row headers for bonus features
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    
    'set a location for each ticker name in a summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'set last row for ticker names
    Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    'set initial start price value of each sheet
    start_Price = ws.Cells(2, 3).Value
    
    'loop through all ticker names
    For i = 2 To Last_Row
    
    'set close price
    Close Price = ws.Cells(i, 6).Value
    
    'set close price
    close_Price = ws.Cells(1, 6).Value
    
    'set open of each new ticker
    Dim New_year_Price As Boolean
    
         If New_year_Price = False Then
         'set opening_Price
         Dim opening_Price As Double
         opening_Price = ws.Cells(i, 3).Value
         
         New_year_Price = True
         
         End If
         
         'check if ticker name is repeating, if not
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
         'set ticker name
         Ticker_Name = ws.Cells(i, 1).Value
         
         'set Yearly Change value
         Yearly_Change = close_Price - opening_Price
         
         'set percent change value
         Percent_change = (Yearly_Change / opening_Price)
         
         'set total volume per ticker
         Total_volume = Total_volume + Cells(i, 7).Value
         
         'add ticker name to summary table
         ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
         
         'add yearly change to Summary table
         ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
         
         'add percent change to Summary table
         ws.Range("L" & Summary_Table_Row).Value = Percentage_Change
         
         'add stock volume per ticker
         ws.Range("M" & Summary_Table_Row).Value = Total_volume
         
         'move to next empty row in summary table
         Summary_Table_Row = Summary_Table_Row + 1
         
         'reset total volume
         Total_volume = 0
         
         'switch back to next price
         New_year_Price = False
         
         Else
         
         Total_volume = Total_volume + ws.Cells(i, 7).Value
         
    End If
    
        'convert yearly change column to show two decimal places and $
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00"
        
        'convert percent change column to show two decimal places and %
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"

    
 Next i
 
        ws.Range("R2") = Increase
        ws.Range("R3") = Decrease
        ws.Range("R4") = GTVolume
        
        'set formatting to show two decimal places and %
        ws.Range("R2:R3").NumberFormat = "0.00%"
        
        'find matching ticker symbols for Max percentage, Min Percentage and MAx Volume
        If ws.Cells(i, 12).Value = Increase Then
        ws.Range("Q2").Value = ws.Cells(i, 10).Value
        
    End If
    
    
        If ws.cell(i, 12).Value = Decrease Then
        ws.Range("Q3").Value = ws.Cells(i, 10).Value
       
    End If
    
        If ws.Cells(i, 13).Value = GTVolume Then
        ws.Range("Q4").Value = ws.Cells(i, 10).Value
        
    End If
    
    
    
    'set last row for summary table
    Final_Row = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    
    For j = 2 To Final_Row
      
      'set conditional formatting to Yearly change column
      If ws.Cells(j, 11).Value > 0 Then
      ws.cell(j, 11).Interior.ColorIndex = 4
      
      Else
      
      ws.Cells(j, 11).Interior.ColorIndex = 3
      
   End If
   
   'set conditional formatting to percent change column
   If ws.Cells(j, 12).Value > 0 Then
   ws.Cells(j, 12).Interior.ColorIndex = 4
   
   Else
   
   ws.Cells(j, 12).Interior.ColorIndex = 3
   
  End If

Next j

  'move to next worksheet
   Next ws
   
    
End Sub



