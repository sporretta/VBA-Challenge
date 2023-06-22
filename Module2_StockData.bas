Attribute VB_Name = "Module1"
'Loop through all worksheets to summarize stock data for each year

Sub StockData()


'Set variables for each worksheet in the workbook

Dim WS_Count As Integer
Dim J As Integer

' Set WS_Count equal to the number of worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop

 For J = 1 To WS_Count
   Worksheets(J).Activate
   
' Set an initial variable for holding the ticker
         
         Dim Ticker As String
         Dim OpenPrice As Double
         Dim ClosePrice As Double
         Dim TotalVolume As Double

'set variable for last row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
'Set variables for summarized data

        TotalVolume = 0
      
' Keep track of the location for each Ticker in the summary table
        Dim StockSummary As Integer
        StockSummary = 2
'Create a new variable to assign the first row of a ticker
    Dim StockRow As Long
    StockRow = 2

' Loop through all Tickers
             For I = 2 To LastRow
  
' Check if we are still within the same ticker, if it is not...
                    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

' Add the values in each ticker category
                    Ticker = Cells(I, 1).Value
                    OpenPrice = Cells(StockRow, 3).Value
                   ClosePrice = Cells(I, 6).Value
                   TotalVolume = TotalVolume + Cells(I, 7).Value
    

' Print the Ticker in the Summary Table
                        Range("I" & StockSummary).Value = Ticker
      
      
' Print the YearlyChange in the Summary Table
                      Range("J" & StockSummary).Value = (ClosePrice - OpenPrice)
      
 ' Print the PercentChange in the Summary Table
                      Range("K" & StockSummary).Value = FormatPercent(((ClosePrice - OpenPrice) / OpenPrice), 2)
      
' Print the Total Volume to the Summary Table

                     Range("L" & StockSummary).Value = TotalVolume
                     
'Color the yearly change green for positive and red for negative
 
           If Range("J" & StockSummary).Value < 0 Then
                Range("J" & StockSummary).Interior.ColorIndex = 3
            
            ElseIf Range("J" & StockSummary).Value > 0 Then
                Range("J" & StockSummary).Interior.ColorIndex = 4
                
           End If
 
' Add one row to the summary table

                    StockSummary = StockSummary + 1
      
' Reset the Prices and Total Volume

                        TotalVolume = 0
                        StockRow = I + 1
' If the cell immediately following a row is the same ticker...
                 Else
                      
                        TotalVolume = TotalVolume + Cells(I, 7).Value

                 End If
              
              
 
        Next I
    
  'Set labels for the greatest % Increase, greatest % Decrease, and greatest Total Volume
  Cells(3, 14).Value = "Greatest % Increase"
  Cells(4, 14).Value = "Greatest % Decrease"
  Cells(5, 14).Value = "Greatest Total Volume"
  Cells(2, 15).Value = "Ticker"
  Cells(2, 16).Value = "Value"
  
  'Define Variables for Max Values
  
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Double
Dim TickerInc As String
Dim TickerDec As String
Dim TickerVol As String
MaxIncrease = 0
MaxDecrease = 0
MaxVolume = 0

  'Use For loop to look for the max values
For I = 2 To LastRow
  
'Find Max Increase %
        
      If Cells(I, 11).Value > MaxIncrease Then
       MaxIncrease = Cells(I, 11).Value
        TickerInc = Cells(I, 9).Value
  
    End If
    
'Find MaxDecrease %

    If Cells(I, 11).Value < MaxDecrease Then
       MaxDecrease = Cells(I, 11).Value
        TickerDec = Cells(I, 9).Value
  
    End If
    
'Find Max Total Volume
   If Cells(I, 12).Value > MaxVolume Then
       MaxVolume = Cells(I, 12).Value
        TickerVol = Cells(I, 9).Value
  
    End If
    
  Next I
  
' Print the Max Values to the Summary Table
         Cells(3, 16).Value = FormatPercent(MaxIncrease, 2)
        Cells(3, 15).Value = TickerInc
        Cells(4, 16).Value = FormatPercent(MaxDecrease, 2)
        Cells(4, 15).Value = TickerDec
        Cells(5, 16).Value = MaxVolume
        Cells(5, 15).Value = TickerVol
      
    Next J
End Sub

