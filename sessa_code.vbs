Attribute VB_Name = "Module1"



Sub Part_1()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
For Each ws In Worksheets
    ws.Activate
    
    'creating end of sheet
    Dim LastRow As Long
    LastRow = Range("A1").CurrentRegion.Rows.Count
    
    'Letting excel know that column A is a string of stock names
    Dim Stock_name As String
    
    'creating counter so it can be summed
    Dim Stock_total As Double
        Stock_total = 0
    
    'creating summary table for where information will go
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'titling table columns
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("M1") = "Open Price"
    Range("N1") = "Close Price"
    
    Dim open_price As Double
    Dim close_price As Double
    
    Dim percent_change As Double
    Dim difference As Double
      
    

        For i = 2 To LastRow

            'checking to see if the ticker names change
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'Setting Ticker Name for each value before it changes to a different value
                Stock_name = Cells(i, 1).Value
                
                'Creating a stock volume sum
                    'creating variable Volume
                
                Stock_total = Stock_total + Cells(i, 7).Value
                
                'close price is the last price value before the change
                close_price = Cells(i, 6).Value
               
            
            'printing the Stock Name, total and price in the summary table
            'looks at the range of I in the Summary Table Row where it equals Stock Name
                Range("I" & Summary_Table_Row).Value = Stock_name
                Range("L" & Summary_Table_Row).Value = Stock_total
                Range("N" & Summary_Table_Row).Value = close_price
              
        
            
            'keep adding additional rows for each new stock name
                Summary_Table_Row = Summary_Table_Row + 1
            
            'Reseting the stock total for the next loop
                Stock_total = 0
            
                
        'Else statement for when the stock name is the same
            Else
                'add to the stock total
                Stock_total = Stock_total + Cells(i, 7).Value
                
                'finding opening price
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    open_price = Cells(i, 3).Value
                   
                    Range("M" & Summary_Table_Row).Value = open_price
                    
                End If
                  
            End If
        'removing zeros so percentage change can be calculated
    Next i
    
 Next ws

End Sub


Sub Part_2()

'creating second loop
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
For Each ws In Worksheets
    ws.Activate

'setting dimensions of second loop variables
Dim LastRow As Long
Dim close_price As Double
Dim open_price As Double


LastRow = Range("A1").CurrentRegion.Rows.Count

   For j = 2 To LastRow
   
     'setting close/open prices to previously made columns
     close_price = Cells(j, 14).Value
     open_price = Cells(j, 13).Value
    
     'if open price is zero setting percentage change to zero to remove division by zero errors
     If open_price = 0 Then
         percentage_change = 0
     Else
                       
         'creating percentage change equation
         difference = close_price - open_price
         percentage_change = difference / open_price
         
         'assigning values to table
         Cells(j, 10).Value = difference
         Cells(j, 11).Value = percentage_change
          
         'assigning color values
         If difference > 0 Then
             Cells(j, 10).Interior.ColorIndex = 4
          Else
             Cells(j, 10).Interior.ColorIndex = 3
                             
         End If
     End If
    Next j
Next

End Sub


