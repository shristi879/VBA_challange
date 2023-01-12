Dim WorksheetName As String
    'Current row
    Dim i As Long
    'Start row of ticker block
    Dim j As Long
    'Index counter to fill Ticker row
    Dim TickCount As Long
    'Last row column A
    Dim LastRowA As Long
    'last row column I
    Dim LastRowI As Long
    'Variable for percent change calculation
    Dim PerChange As Double
    'Variable for greatest increase calculation
    Dim GreatIncr As Double
    'Variable for greatest decrease calculation
    Dim GreatDecr As Double
    'Variable for greatest total volume
    Dim GreatVol As Double
    
    'Get the WorksheetName
    WorksheetName = Ws.Name
    
    'Create column headers
    Ws.Cells(1, 9).Value = "Ticker"
    Ws.Cells(1, 10).Value = "Yearly Change"
    Ws.Cells(1, 11).Value = "Percent Change"
    Ws.Cells(1, 12).Value = "Total Stock Volume"
    Ws.Cells(1, 16).Value = "Ticker"
    Ws.Cells(1, 17).Value = "Value"
    Ws.Cells(2, 15).Value = "Greatest % Increase"
    Ws.Cells(3, 15).Value = "Greatest % Decrease"
    Ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Set Ticker Counter to first row
    TickCount = 2
    
    'Set start row to 2
    j = 2
    
    'Find the last non-blank cell in column A
    LastRowA = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox ("Last row in column A is " & LastRowA)
    
        'Loop through all rows
        For i = 2 To LastRowA
        
            'Check if ticker name changed
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
            'Write ticker in column I (#9)
            Ws.Cells(TickCount, 9).Value = Ws.Cells(i, 1).Value
            
            'Calculate and write Yearly Change in column J (#10)
            Ws.Cells(TickCount, 10).Value = Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value
            
                'Conditional formating
                If Ws.Cells(TickCount, 10).Value < 0 Then
            
                'Set cell background color to red
                Ws.Cells(TickCount, 10).Interior.ColorIndex = 3
            
                Else
            
                'Set cell background color to green
                Ws.Cells(TickCount, 10).Interior.ColorIndex = 4
            
                End If
                
                'Calculate and write percent change in column K (#11)
                If Ws.Cells(j, 3).Value <> 0 Then
                PerChange = ((Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value) / Ws.Cells(j, 3).Value)
                
                'Percent formating
                Ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                
                Else
                
                Ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                
                End If
                
            'Calculate and write total volume in column L (#12)
            Ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(Ws.Cells(j, 7), Ws.Cells(i, 7)))
            
            'Increase TickCount by 1
            TickCount = TickCount + 1
            
            'Set new start row of the ticker block
            j = i + 1
            
            End If
        
        Next i
        
    'Find last non-blank cell in column I
    LastRowI = Ws.Cells(Rows.Count, 9).End(xlUp).Row
    'MsgBox ("Last row in column I is " & LastRowI)
    
    'Prepare for summary
    GreatVol = Ws.Cells(2, 12).Value
    GreatIncr = Ws.Cells(2, 11).Value
    GreatDecr = Ws.Cells(2, 11).Value
    
        'Loop for summary
        For i = 2 To LastRowI
        
            'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
            If Ws.Cells(i, 12).Value > GreatVol Then
            GreatVol = Ws.Cells(i, 12).Value
            Ws.Cells(4, 16).Value = Ws.Cells(i, 9).Value
            
            Else
            
            GreatVol = GreatVol
            
            End If
            
            'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
            If Ws.Cells(i, 11).Value > GreatIncr Then
            GreatIncr = Ws.Cells(i, 11).Value
            Ws.Cells(2, 16).Value = Ws.Cells(i, 9).Value
            
            Else
            
            GreatIncr = GreatIncr
            
            End If
            
            'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
            If Ws.Cells(i, 11).Value < GreatDecr Then
            GreatDecr = Ws.Cells(i, 11).Value
            Ws.Cells(3, 16).Value = Ws.Cells(i, 9).Value
            
            Else
            
            GreatDecr = GreatDecr
            
            End If
            
        'Write summary results in ws.Cells
        Ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
        Ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
        Ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
        
        Next i
        
    'Djust column width automatically
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
       
Next Ws

End Sub
