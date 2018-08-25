Sub credit_card()
 
   Dim ws As Worksheet
  ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets


        ' Set an initial variable for holding the ticker name
            Dim ticker_Name As String
        
            ' Set an initial variable for holding the total Volume
            Dim volume_Total As Double
            volume_Total = 0
        
            ' Keep track of the location for each ticker name in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
            'creating this variable to count the total number of rows before reaching the last observation for each ticker use to calculate the difference between close and open price
            Dim tick_n As Integer
            
            'Initially set it as 0 
            tick_n = 0
   
            Dim volumemax As Integer
           
             ' Determine the Last Row in eahc sheet
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            ' Loop through all tickers
            For i = 2 To lastrow
        
                  ' Check if we are still within the same credit card brand, if it is not...
                  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    ' Set the ticker name
                    ticker_Name = ws.Cells(i, 1).Value
            
                    ' Add to the Volume total
                    volume_Total = volume_Total + ws.Cells(i, 7).Value
            
                    ' Print the ticker name in the Summary Table
                    ws.Range("i" & Summary_Table_Row).Value = ticker_Name
            
                    ' Print the total volume to the Summary Table
                    ws.Range("j" & Summary_Table_Row).Value = volume_Total
                    
                    'closing price
                     close_price = ws.Cells(i, 6).Value
                    'opening price by substracting the number of rows before the last obs for each ticker
                     open_price = ws.Cells(i - tick_n, 3).Value
                     
                     'Calculating the difference 
                     ws.Range("k" & Summary_Table_Row).Value = close_price - open_price
                     
                     'Assigning colors to each cell 
                     If ws.Range("k" & Summary_Table_Row).Value >= 0 Then
                           
                            ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
                     Else
                            ws.Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                     End If
                     
                     'Taking 0 into consideration
                        If open_price <> 0 Then
                        
                        ws.Range("l" & Summary_Table_Row).Value = ((close_price - open_price) / open_price) * 100 & "%"
            
                        Else
                        
                        ws.Range("l" & Summary_Table_Row).Value = "NA"
                        
                        End If
            
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    ' Reset the Volume total 
                    volume_Total = 0
                    
                    tick_n = 0
            
                
                  Else
                    ' Reset the Volume total
                    volume_Total = volume_Total + ws.Cells(i, 7).Value
            
                   'accruting the tick_n
                    tick_n = tick_n + 1
                    
                  End If
                    ' Add one to the summary table row
        
            Next i
    
         
          'Initializing the maximum value for total volume and maixmum and minimum for percent change by using the first observation
            Maxpercent = ws.Cells(2, 12).Value
            Minpercent = ws.Cells(2, 12).Value
            Maxvolume = ws.Cells(2, 10).Value
            Maxpercent_NAME = ws.Cells(2, 9).Value
            Minpercent_NAme = ws.Cells(2, 9).Value
            Maxvolume_name = ws.Cells(2, 9).Value
            
         'For loop to Identifying maximum/minimum number and associated ticker name
            For j = 3 To (Summary_Table_Row - 1)
            
                If ws.Cells(j, 10).Value > Maxvolume Then
                  Maxvolume = ws.Cells(j, 10).Value
                  Maxvolume_name = ws.Cells(j, 9)
                Else
                  Maxvolume = Maxvolume
                End If
               
                If ws.Cells(j, 12).Value > Maxpercent And ws.Cells(j, 12) <> "NA" Then
                  Maxpercent = ws.Cells(j, 12).Value
                  Maxpercent_NAME = ws.Cells(j, 9)
                Else
                  Maxpercent = Maxpercent
                End If
               
                If ws.Cells(j, 12).Value < Minpercent And ws.Cells(j, 12) <> "NA" Then
                  Minpercent = ws.Cells(j, 12).Value
                  Minpercent_NAme = ws.Cells(j, 9)
                Else
                Minpercent = Minpercent
                End If
            
            Next j
                
          'Assigning names in the worksheet
            ws.Range("P2").Value = Maxpercent * 100 & "%"
            ws.Range("P3").Value = Minpercent * 100 & "%"
            ws.Range("P4").Value = Maxvolume
            ws.Range("O2").Value = Maxpercent_NAME
            ws.Range("O3").Value = Minpercent_NAme
            ws.Range("O4").Value = Maxvolume_name
            ws.Range("N2").Value = "Greatest increase %"
            ws.Range("N3").Value = "Greatest decrease %"
            ws.Range("N4").Value = "Greatest Volume"
            
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Total Volume"
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L1").Value = "Percent Change"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Percent Change"
                 
       Next ws

End Sub


