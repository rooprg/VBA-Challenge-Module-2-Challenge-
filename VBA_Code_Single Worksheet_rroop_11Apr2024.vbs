Sub RoopStockAnaylsisSingle()

'Part 1
'Step 1: Setting Definitions
Dim ticker As String
Dim Volume As Double
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Summary_Table_Row As Double

'Step 2: Setting Outputs and Locations
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Range("I1").Value = "ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("I1:L1").Font.Bold = True
Range("P1:Q1").Font.Bold = True

Summary_Table_Row = 2

'Step 3: Looping through stock data
For i = 2 To UsedRange.Rows.Count
   If i = 2 Then
    Year_Open = Cells(i, 3).Value
    End If

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ticker = Cells(i, 1).Value
    Volume = Volume + Cells(i, 7)
    Year_Close = Cells(i, 6)
    Yearly_Change = Year_Close - Year_Open
    Percent_Change = ((Year_Close - Year_Open) / Year_Open)
    
    Cells(Summary_Table_Row, 9).Value = ticker
    Cells(Summary_Table_Row, 10).Value = Yearly_Change
    Cells(Summary_Table_Row, 11).Value = Percent_Change
    Cells(Summary_Table_Row, 12).Value = Volume
    Summary_Table_Row = Summary_Table_Row + 1
    
    Year_Open = Cells(i + 1, 3).Value
    Year_Close = 0
    Volume = 0
    
    Else
    Volume = Volume + Cells(i, 7).Value
    
    End If
    
Next i

'Step 4: Add formatting to Columns
'Step 4a - Adding % to Percent Change
    Columns("K").NumberFormat = "0.00%"
    
'Step 4b - Shading cells equal/over and under 0
    For j = 2 To (Cells(Rows.Count, 10).End(xlUp).Row)
        If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        Cells(j, 11).Interior.ColorIndex = 4
        
        Else
        
        Cells(j, 10).Interior.ColorIndex = 3
        Cells(j, 11).Interior.ColorIndex = 3
        
        End If
        
    Next j
    
'Part 2
'Step 1: Tickers and Values
'Step 1a: Defining Tickers and Values for Greatests
 Dim Ticker_GPI As String
 Dim Ticker_GDI As String
 Dim Ticker_GTV As String
 Dim Greatest_Percent_Increase As Double
 Dim Greatest_Percent_Decrease As Double
 Dim Greatest_Total_Volume As Double
 
 'Step 1b: Setting initial values
 Greatest_Percent_Increase = 0
 Greatest_Percent_Decrease = 0
 Greatest_Total_Volume = 0
 
 'Step 1c: Looping through to retrieve Tickers and Values for Greatests
    For k = 2 To Cells(Rows.Count, 10).End(xlUp).Row
        
        If Cells(k, 11).Value > Greatest_Percent_Increase Then
            Greatest_Percent_Increase = Cells(k, 11).Value
            Ticker_GPI = Cells(k, 9).Value
        
        End If
        
        If Cells(k, 11).Value < Greatest_Percent_Decrease Then
            Greatest_Percent_Decrease = Cells(k, 11).Value
            Ticker_GDI = Cells(k, 9).Value
            
        End If
        
        If Cells(k, 12).Value > Greatest_Total_Volume Then
            Greatest_Total_Volume = Cells(k, 12).Value
            Ticker_GTV = Cells(k, 9).Value
            
        End If
        
    Next k
  
  'Step 2: Outputing tickers and values to table locations
  Cells(2, 16) = Ticker_GPI
  Cells(3, 16) = Ticker_GDI
  Cells(4, 16) = Ticker_GTV
  
  Cells(2, 17) = Greatest_Percent_Increase
  Cells(3, 17) = Greatest_Percent_Decrease
  Cells(4, 17) = Greatest_Total_Volume
  
  'Step 3: Formatting Values
  Cells(2, 17).NumberFormat = "0.00%"
  Cells(3, 17).NumberFormat = "0.00%"

 
End Sub
