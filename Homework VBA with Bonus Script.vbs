Attribute VB_Name = "CompilesToEachTab"




'VBA - HOMEWORK #2
'THERE ARE TWO VERSIONS OF THIS SCRIPT
'       SummarizeStockMarketData_Compile_On_EACH_Tab() COMPILES SUMMARY DATA FOR THE A TAB ON THE SAME TAB
'       Sub SummarizeStockMarketData_CompileToOneTab() COMPILES SUMMARY DATA FOR EACH TAB ON A SINGLE TAB - THE FIRST IN THE COLLECTION OF WORKSHEETS


'----------------------------------------------------------------------------------------------------------------

'Key Assumption: Data is sorted by ticker and in chronological order

'----------------------------------------------------------------------------------------------------------------

Sub SummarizeStockMarketData_Compile_On_EACH_Tab()

 'Worksheet variables
 Dim WS_Count As Integer
 Dim WS_I As Integer
 

' Set a variable for specifying the column of interest
Dim Column As Integer
Dim Row As Integer
Dim LastRow As Variant
Dim TickerRowCount
Dim CurrentOne As String
Dim OneBelow As String
Dim CaptureCounter As Integer
Dim OpeningPrice As Variant
Dim ClosingPrice As Variant
Dim VolumeCounter As Variant
Dim PercentChange As Variant
Dim RowCounter As Variant

' Set WS_Count equal to the number of worksheets in the active workbook.
 WS_Count = ActiveWorkbook.Worksheets.Count


' Begin the loop that creates the summary data in each tab
For WS_I = 1 To WS_Count
            
           'MsgBox ActiveWorkbook.Worksheets(I).Name
            Worksheets(WS_I).Activate
            
           
            'Column Headers for each summary datapoint
           Cells(1, 9).Value = "Ticker"
           Cells(1, 10).Value = "Yearly Change"
           Cells(1, 11).Value = "Percent Change"
           Cells(1, 12).Value = "Total Stock Volumne"

        'Initializes the row for writing captured summary data for each tab
        CaptureCounter = 2



        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    
            
            'initializes total for volume running total
            VolumeCounter = 0
            
            'Initializes the first row (the first date) for for the ticker
            TickerRowCount = 1
          
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
         
          
                  
        'READS AND WRITES STOCK MARKET DATA ON EACH TAB AND CONSOLIDATES ON AN AREA IN THE FIRST TAB
                  
                  
                ' Loop through ticker rows and reads ticker, opening price, closing price and volume
                 For I = 2 To LastRow
                        
                        
                        'activate the current tab for summary data capture
                        Worksheets(WS_I).Activate
                        
                        'READ CURRENT TICKER
                        CurrentOne = Cells(I, 1).Value
                        
                        'READ NEXT TICKER TO MARK TRANSITION
                        OneBelow = Cells(I + 1, 1).Value
                        
                        'READ VOLUME
                        Volume = Cells(I, 7).Value
                        'Running total for Volume
                        VolumeCounter = VolumeCounter + Volume
                        
                        
                        'READ OPENING PRICE
                        If TickerRowCount = 1 Then
                        
                        'Capture Opening Price on the first day and save for later use
                            'If the opening price is 0 it will create a divide by zero error.
                                'If the opening price is a zero, then take the low for that day
                                                If OpeningPrice = 0 Then
                                                    '<Low>
                                                    OpeningPrice = Cells(I, 5).Value
                                                    
                                                    '<Opening Price>
                                                    Else: OpeningPrice = Cells(I, 3).Value
                                                
                                                End If
                         End If
                                       
                        TickerRowCount = TickerRowCount + 1
                
                        ' READS CLOSING PRICE\
                        '   Tests for change in tickers to identify change
                        If CurrentOne <> OneBelow Then
                        
                                    'Read Closing Price and save for later use
                                    ClosingPrice = Cells(I, 6).Value
                                                                                            
                                    
                    ' WRITES DATA TO STAGING AREA
                                    'Activate the tab for writing to the staging area
                                    'Worksheets(I).Activate
                                    
                                    'WRITE TICKER
                                    'Cells(CaptureCounter, 9).Select
                                    Cells(CaptureCounter, 9).Value = CurrentOne
                                    
                                    'Write Opening Price
                                    'Cells(CaptureCounter, 15).Value = OpeningPrice
                                                                
                                    'Write Closing Price
                                    'Cells(CaptureCounter, 16).Value = ClosingPrice
                                    
                                    'WRITE PRICE CHANGE
                                    YearlyChange = ClosingPrice - OpeningPrice
                                    Cells(CaptureCounter, 10).Value = YearlyChange
                                        'FORMAT YEARLY CHANGE
                                        If YearlyChange < 0 Then
                                            Cells(CaptureCounter, 10).Interior.ColorIndex = 3
                                            
                                            Else
                                            Cells(CaptureCounter, 10).Interior.ColorIndex = 4
                                    
                                        End If
                                                        
                                    'WRTIE PERCENT CHANGE
                                    PercentChange = (ClosingPrice / OpeningPrice - 1)
                                    
                                    'format percent
                                    Cells(CaptureCounter, 11).Value = PercentChange
                                    'Cells(CaptureCounter, 11).Style = "Percent"
                                    Cells(CaptureCounter, 11).NumberFormat = "0.00%"
                                    
                                    'WRITE VOLUME
                                    Cells(CaptureCounter, 12).Value = VolumeCounter
                                    Cells(CaptureCounter, 12).NumberFormat = "0,000"
                
                                    'Reinitialize the Row Counter for the Symbol
                                    
                                    TickerRowCount = 1
                                    
                                    'Reinitialize the Volume Counter for the Symbol
                                    VolumeCounter = 0
                                            
                                    ' Message Box the value of the current cell and value of the next cell
                                    'MsgBox (CurrentOne & " and then " & OneBelow)
                                    
                                    'Increment the CaptureCounter by 1
                                    CaptureCounter = CaptureCounter + 1
                
                        End If
        
        
                        
        
        'Go to the next row
                  Next I

           
        
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
 
                     
'Go to the next tab
Next WS_I
    
Call Bonus
    
    
Range("a1").Select
     
End Sub

Sub Bonus()

 'Worksheet variables
Dim WS_Count As Integer
Dim WS_I As Integer

' Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

For WS_I = 1 To WS_Count


            
           'MsgBox ActiveWorkbook.Worksheets(I).Name
            Worksheets(WS_I).Activate
            
           
            'Column Headers for each bonus summary datapoint
           Cells(1, 15).Value = "Ticker"
           Cells(1, 16).Value = "Value"
           
           
            'Row headers
            Cells(2, 14).Value = "Greatest % Increase"
            Cells(3, 14).Value = "Greatest % Decrease"
            Cells(4, 14).Value = "Greates Total Volume"
            
            
            'Tickers
            Cells(2, 15).FormulaR1C1 = "=INDEX(R1C9:R9001C12,MATCH(RC[1],C[-4],0),1)"
            Cells(3, 15).FormulaR1C1 = "=INDEX(R1C9:R9001C12,MATCH(RC[1],C[-4],0),1)"
            Cells(4, 15).FormulaR1C1 = "=INDEX(R1C9:R9001C12,MATCH(RC[1],C[-3],0),1)"
            
            'Values
            Cells(2, 16).FormulaR1C1 = "=MAX(R1C11:R9001C11)"
            Cells(2, 16).NumberFormat = "0.00%"
            Cells(3, 16).FormulaR1C1 = "=MIN(R1C11:R9001C11)"
            Cells(3, 16).NumberFormat = "0.00%"
            Cells(4, 16).FormulaR1C1 = "=Max(R1C12:R9001C12)"
            NumberFormat = "0,000"
           
            'Autofit all columns
            Cells.Select
            Cells.EntireColumn.AutoFit
            
            'Select home
            Range("a1").Select
 
 
 
 'Go to the next tab
Next WS_I
              
              

Worksheets(1).Activate
     

End Sub




'
'        End If

