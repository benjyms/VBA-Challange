Attribute VB_Name = "Ticker_Tracking"
'References include Microsoft Macro Loop

Sub ticker_work()

    ' _____________________
    '
    ' LOOP THROUHG ALL SHEETS
    '______________________
    
    
    
    'Declare the as a worksheet object variable
    
    Dim Current As Worksheet
    Dim LastRow As Long
    
    
       
    'Loop through all worksheets in the active workbook
    
            
        For Each Current In Worksheets
                                
                'Add "Ticker" Header
                
                Current.Range("I1").Value = ("Ticker")
                
                'Add "Yearly Change" Header
                
                Current.Range("J1").Value = ("Yearly Change")
                
                'Add "Percent Change" Header
                
                Current.Range("K1").Value = ("Percent Change")
                
                'Add "Total Stock Volume" Header
                
                Current.Range("L1").Value = ("Total Stock Volume")
                
                    
                                
                    'BEGINNING THE LOOP CAPTURE EACH TICKERS yR cHG, % CHN, AND TOTAL VOLUME
                    
                    Dim Ticker As String
                    Dim Yr_change As Double
                    Dim Per_change As Double
                    Dim Open_price As Double
                    Dim Close_price As Double
                    Dim Total_volume As Double
                                        
                    LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row
                    
                    'Keep track of the location in the Sumamry table to place the next Ticker Symbol
                    
                    Dim Summary_Table_Row As Integer
                    Summary_Table_Row = 2
                    Total_volume = 0
                    Open_price = 0
                    Close_price = 0
                    
                    
                    
                'Loop through all ticker rows on the
                    
            For I = 2 To LastRow
                
                    'Capture the Open price from first record of ticker
                    If Open_price = 0 Then
                    
                        Open_price = Current.Cells(I, 3).Value
                        
                        Else
                    
                        Open_price = Open_price
                        
                    End If
                                            
                    
            If Current.Cells(I + 1, 1).Value <> Current.Cells(I, 1).Value Then
                            
                    'Setting the Summary Ticker Value
                    Ticker = Current.Cells(I, 1).Value
                    
                    'Capture the Close Price
                    Close_price = Current.Cells(I, 6).Value
                    
                    'Setting Total Volume
                    Total_volume = Total_volume + Current.Cells(I, 7)
                    
                    'Print the Ticker Symbol in the Summary Table
                    Current.Range("I" & Summary_Table_Row).Value = Ticker
                    
                    'Print the total volume in the Summary Table
                    Current.Range("L" & Summary_Table_Row).Value = Total_volume
                                                
                    'Print the yearly change
                    Yr_change = Close_price - Open_price
                    
                    Current.Range("J" & Summary_Table_Row).Value = Yr_change
                        
                        'conditinoal formatting for yearly change
                                                                                       
                        If Current.Range("J" & Summary_Table_Row).Value > 0 Then
                        
                        Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                        ElseIf Current.Range("J" & Summary_Table_Row).Value < 0 Then
                                                 
                        Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                        Else
                        
                        Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 2
                        
                        End If
                                                        
                                              
                    'calculate and print the yearly percentage change and account for 0s
                    If Yr_change = 0 And Close_price = 0 Then
                    
                    Per_change = 0
                    
                    Else
                    
                    Per_change = Yr_change / Close_price
                    
                    End If
                        
                        
                                                                                             
                    Current.Range("K" & Summary_Table_Row).Value = Per_change
                    
                    'change result to percentage
                    
                    Current.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                                                
                    'Advance the Summary Table Row by 1 for next ticker
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    'Reset Total_volume to begin adding for next ticker
                    Total_volume = 0
                    
                    'Reset Open_price
                    Open_price = 0
                    
                    'Reset Close_price
                    Close_price = 0
                    
                    'Reset the Yr_change
                    Yr_change = 0
                    
                    
                            
            'If the cell immediately following is the same ticker
                    
            Else
            
                Total_volume = Total_volume + Current.Cells(I, 7).Value
            
            End If
                        
        Next I
        
    Current.Cells.Columns.AutoFit
                           
    Next
        
            
End Sub

