Attribute VB_Name = "Module1"
Sub stock_data()
    
    'Labeling headers in column's "I, J, K, L"
For Each ws In Worksheets

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"


'Assigning variable dimensions
Dim Summary_Table As Integer
Dim Ticker_Symbol As String
Dim opn As Double
Dim cls As Double
Dim countItem As Double

Dim Yearly_Change As Double
   
  'Assigning numerical values to variable dimensions
   Summary_Table = 2
   Yearly_Change = 0
   opn = 0
   cls = 0
   countItem = 0



'The "Total Rows" is all the cells in column A stopping at the last cell in the data set
Total_Rows = ws.Cells(Rows.Count, "A").End(xlUp).Row


'Pulling tickers into Column I
    For i = 1 To Total_Rows

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ws.Cells(Summary_Table, 9).Value = ws.Cells(i + 1, 1).Value
            
            'Establishing first open value for each ticker
            opn = ws.Cells(i + 1, 3).Value
            
            countItem = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(Summary_Table, 9).Value)
            
            'Establishing last close value for each ticket
            cls = ws.Cells(i + countItem, 6).Value
            
            'Calculating yearly change by Subtracting open from close
            ws.Cells(Summary_Table, 10).Value = cls - opn
            
            'Formatting numbers
            ws.Cells(Summary_Table, 10).NumberFormat = "0.00"
            
               
            'Percent change setup and math
            If opn = 0 Then
                ws.Cells(Summary_Table, 11).Value = 0
            Else
                ws.Cells(Summary_Table, 11).Value = (cls - opn) / opn
                                 
            End If
              
                          
            Summary_Table = Summary_Table + 1
        
        End If
                                  
    
    Next i
        
        
    ticker_report = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    
    'Establishing the color of the cell for positive and negative value
    For cnt = 2 To ticker_report
      
      
      If ws.Cells(cnt, 10).Value >= 0 Then
      
        ws.Cells(cnt, 10).Interior.ColorIndex = 4
        
      Else
      
        ws.Cells(cnt, 10).Interior.ColorIndex = 3
      
      End If
      
'Assigning ranges to setup up SumIfs function to sum cells that match multiple criteria
 Dim arange As Range
 Dim crange As Range
 Dim grange As Range
  
    Set arange = ws.Range("A:A")
    Set crange = ws.Range("C:C")
    Set grange = ws.Range("G:G")
      
      
      
       ws.Cells(cnt, 12).Value = WorksheetFunction.SumIfs(grange, arange, ws.Cells(cnt, 9).Value)
         
         
Next cnt
    
    
    ws.Columns("K").NumberFormat = "0.00%"
    
    
    'Creating labels for the Greatest % increase, greatest % decrease, and greatest total volume
        ws.Cells(2, 16).Value = "Greatest % Increase"
    
        ws.Cells(3, 16).Value = "Greatest % Decrease"
    
        ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    
    'Creating headers to hold Ticker symbol and value for the different labels
    ws.Cells(1, 17).Value = "Ticker"
    
    ws.Cells(1, 18).Value = "Value"
   
    
    
    ws.Cells(2, 18).Value = WorksheetFunction.Max(ws.Range("K:K"))
     
    ws.Cells(3, 18).Value = WorksheetFunction.Min(ws.Range("K:K"))
     
    ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("L:L"))
     
     

    For cnt3 = 2 To ticker_report
        'Matching the "Greatest % Increase" with the corresponding Ticker Symbol
        If ws.Cells(cnt3, 11).Value = ws.Cells(2, 18).Value Then
            
            ws.Cells(2, 17).Value = ws.Cells(cnt3, 9).Value
        
        'Matching the "Least % Increase" with the corresponding Ticker Symbol
        ElseIf ws.Cells(cnt3, 11).Value = ws.Cells(3, 18).Value Then

            ws.Cells(3, 17).Value = ws.Cells(cnt3, 9).Value

        ElseIf ws.Cells(cnt3, 12).Value = ws.Cells(4, 18).Value Then

            ws.Cells(4, 17).Value = ws.Cells(cnt3, 9).Value

        End If

     Next cnt3
     
     ws.Cells(2, 18).NumberFormat = "0.00%"
     ws.Cells(3, 18).NumberFormat = "0.00%"
     ws.Cells(4, 18).NumberFormat = "0.00"
         
    
Next ws


End Sub
  
