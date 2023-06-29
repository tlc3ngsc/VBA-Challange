Sub WarrenBuffet()
' --- Section 01
' ----------- Define Variabls ---------------------------------------------------------------------
Dim I As Integer
Dim K As Integer
Dim StockPriceEndRowColumn As Double
Dim StockPriceStart As Double
Dim StockPriceExit As Double
Dim StockPriceEnd As Double
Dim StockPriceCounter As Integer
Dim StockTicker As String
Dim CurrentCell As String
Dim CurrentCellPlus1 As String
Dim CurrentStockPricePlus1 As Double
Dim CurrentStockPrice As Double
Dim Transactions As Double
Dim StockPriceLow As Double
Dim StockPriceHigh As Double
Dim StockPriceStartRowColumn As Double
Dim StockPriceDifference As Double
Dim StockPricePercentage As Double
Dim StockPricePercentageRowColumn As Integer
Dim StockPriceDifferenceRowColumn As Double
Dim StockPriceHighRowColumn As Integer
Dim StockPriceLowColumn As Integer
Dim PercentChange As Double
Dim StartRow As Double
Dim StockPriceColumn As Integer
Dim LastRow As Double
Dim NextRow As Double
Dim StartColumn As Integer
Dim TransactionColumn As Integer
Dim TransactionRow As Integer
Dim TransactionsRowColumn As Double
Dim TranactionsTotal As Double
Dim TickerSys As String
Dim TickerYrChng As Double
Dim TickerPerChng As Double
Dim TickerVol As Double
Dim TickerRowPlace As Integer
Dim TickerRowColumn As Integer
Dim CheckValues As Double
Dim GreatestTotalVolume As Double
Dim GreatesPercentIncrease As Double
Dim GreatesPercentDecrease As Double
Dim TickerSysmbolGreatestTotalVolume As String
Dim TickerSysmbolGreatesPercentIncrease As String
Dim TickerSysmbolGreatesPercentDecrease As String
Dim ws As Worksheet

' ----------------------------------------------------------------
For Each ws In Worksheets
    I = 2
    StartColumn = 1
    TickerRowPlace = 1
    StockPriceColumn = StartColumn + 5
    StockPriceHigh = ws.Cells(I, StockPriceColumn).Value
    TransactionRow = 2
    TransactionColumn = 7
    TickerRowColumn = 9
    CurrentCell = ws.Cells(I, StartColumn).Value
    NextRow = I + 1
    CurrentCellPlus1 = ws.Cells(NextRow, StartColumn).Value
    LastRow = ws.Range("A1").End(xlDown).Row
    ' ---------------------------------------------------------------------------------------------------------------
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Transactions"
    ws.Range("K1").Value = "High Stock Price"
    ws.Range("L1").Value = "High Stock Low"
    ws.Range("M1").Value = "Stock Price Begin Price verse Ending Price Difference"
    ws.Range("N1").Value = "StockPricePercentage"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Percentage"
    ws.Range("O2").Value = "GreatesPercentIncrease"
    ws.Range("O3").Value = "GreatesPercentDecrease"
    ws.Range("O4").Value = "Greatest Total Volume"
'  ws.Range("Q3").NumberFormat = "0.0%"
  ws.Range("Q2").NumberFormat = "0.0%"
    GreatesPercentIncrease = 0
    GreatesPercentDecrease = 0
    StockPriceLow = 100000000
    StartRow = 0
    StockPriceHigh = 0
    Transactions = 0
   GreatestTotalVolume = 0
                   
    K = 0
    StockPriceCounter = 0
    StockPriceStartRowColumn = 3
       
' ---------------------------------------------------------------------------------------------------------------
    For StartRow = I To LastRow
        CurrentCell = ws.Cells(StartRow, StartColumn).Value
        CurrentCellPlus1 = ws.Cells(NextRow, StartColumn).Value
        If StockPriceCounter = 0 Then
               StockPriceStart = ws.Cells(StartRow, StockPriceStartRowColumn).Value
               StockPriceCounter = 1
         End If
         
        
          If CurrentCell = CurrentCellPlus1 Then
               CurrentStockPrice = ws.Cells(StartRow, StockPriceColumn).Value
               CurrentStockPricePlus1 = ws.Cells(NextRow, StockPriceColumn).Value
                CheckValues = ws.Cells(StartRow, 7).Value
                Transactions = Transactions + ws.Cells(StartRow, 7).Value
                 If StockPriceHigh < CurrentStockPrice Then
                    StockPriceHigh = CurrentStockPrice
                 End If
                 If StockPriceLow > CurrentStockPrice Then
                    StockPriceLow = CurrentStockPrice
                 End If
                 If GreatesPercentIncrease < StockPricePercentage Then
                   GreatesPercentIncrease = StockPricePercentage
                ws.Range("Q2").Value = GreatesPercentIncrease
               ws.Range("Q2").NumberFormat = "0.0%"
                   ws.Range("P2").Value = CurrentCell
             Else
                  ws.Range("Q2").Value = GreatesPercentIncrease
                   ws.Range("Q2").NumberFormat = "0.0%"
                  ws.Range("P2").Value = CurrentCell
           
              End If
            

          Else
               StockPriceEndRowColumn = StartRow
               StockPriceExit = ws.Cells(StockPriceEndRowColumn, 6).Value
               StockPriceDifference = StockPriceExit - StockPriceStart
   
               
   ' ------------------------------------------------------------------------------------------------------------
              If StockPriceStart = StockPriceExit Then
                  StockPricePercentage = 0
              Else
                 StockPricePercentage = ((StockPriceExit - StockPriceStart) / StockPriceStart)
              End If
              If StockPriceHigh < StockPriceExit Then
                 StockPriceHigh = StockPriceExit
              End If
              If StockPriceLow > StockPriceExit Then
                 StockPriceLow = StockPriceExit
              End If
 ' ------------------------------------------------------------------------------------------------------------
             
              Transactions = Transactions + ws.Cells(StartRow, 7).Value
              TickerSys = CurrentCell
              TickerRowPlace = TickerRowPlace + 1
              TransactionRowColumn = TickerRowColumn + 1
              StockPriceHighRowColumn = TickerRowColumn + 2
              StockPriceLowRowColumn = TickerRowColumn + 3
              StockPriceDifferenceRowColumn = TickerRowColumn + 4
              StockPricePercentageRowColumn = TickerRowColumn + 5
                      
' ------------------------------------------------------------------------------------------------------------
              ws.Cells(TickerRowPlace, StockPriceHighRowColumn).Value = StockPriceHigh
              ws.Cells(TickerRowPlace, StockPriceLowRowColumn).Value = StockPriceLow
              ws.Cells(TickerRowPlace, TickerRowColumn).Value = CurrentCell
              ws.Cells(TickerRowPlace, TransactionRowColumn).Value = Transactions
              ws.Cells(TickerRowPlace, StockPriceDifferenceRowColumn).Value = StockPriceDifference
              ws.Cells(TickerRowPlace, StockPricePercentageRowColumn).Value = StockPricePercentage
              ws.Cells(TickerRowPlace, StockPricePercentageRowColumn).NumberFormat = "0.0%"
' --------------------------------------------------------------------------------------------------Column 0 = 15 Column P = 16 Column Q = 17 ----
            If GreatesPercentIncrease < StockPricePercentage Then
                   GreatesPercentIncrease = StockPricePercentage
                   ws.Range("Q2").Value = GreatesPercentIncrease
                   ws.Range("Q2").NumberFormat = "0.0%"
                   ws.Range("P2").Value = CurrentCell
             Else
                     ws.Range("Q2").Value = GreatesPercentIncrease
                  ws.Range("Q2").NumberFormat = "0.0%"
                  ws.Range("P2").Value = CurrentCell
           
             End If
 
'       ----------------------------------------------------------------------------------------------------------
              StockPriceLow = 10000000
              StockPricePercentage = 0
              StockPriceHigh = 0
              Transactions = 0
              K = 0
              StockPriceStart = 0
              StockPriceCounter = 0
        End If
        NextRow = NextRow + 1
    Next StartRow
 ' ------------------------------------------------------------------------------------------------------------
   GreatesPercentIncrease = 0
   TickerSysmbolGreatesPercentIncrease = " "
   GreatesPercentDecrease = 0
   GreatestTotalVolume = 0
   StartRow = 0
   LastRow = ws.Range("J1").End(xlDown).Row
   For K = 2 To LastRow
     If ws.Cells(K, 10).Value >= GreatestTotalVolume Then
         GreatestTotalVolume = ws.Cells(K, 10).Value
         StartRow = K
     End If
   Next K
   ws.Range("Q4").Value = GreatestTotalVolume
   ws.Range("P4").Value = ws.Cells(StartRow, 9).Value
' ------------------------------------------------------------------------------------------------------------
   StartRow = 0
   K = 0
   GreatesPercentIncrease = 0
   LastRow = ws.Range("N1").End(xlDown).Row
   For K = 2 To LastRow
      If ws.Cells(K, 14).Value >= GreatesPercentIncrease Then
          GreatesPercentIncrease = ws.Cells(K, 14).Value
          StartRow = K
      End If
   Next K
   ws.Range("Q2").Value = GreatesPercentIncrease
   ws.Range("Q2").NumberFormat = "0.0%"
   ws.Range("P2").Value = ws.Cells(StartRow, 9).Value
' ------------------------------------------------------------------------------------------------------------
   StartRow = 0
   I = 0
   CheckValues = 0
   GreatesPercentDecrease = -0.000001
   LastRow = ws.Range("N1").End(xlDown).Row
   For I = 2 To LastRow
   CheckValues = ws.Cells(I, 14).Value
      If ws.Cells(I, 14).Value <= GreatesPercentDecrease Then
          GreatesPercentDecrease = ws.Cells(I, 14).Value
          StartRow = I
      End If
   Next I
   LastRow = ws.Range("N1").End(xlDown).Row
   I = 0
    For I = 2 To LastRow
       If ws.Cells(I, 14).Value < 0 Then
          ws.Cells(I, 14).Interior.ColorIndex = 3
          StartRow = I
          Else
          ws.Cells(I, 14).Interior.ColorIndex = 5
      End If
   Next I
   
   ws.Range("Q3").Value = GreatesPercentDecrease
   ws.Range("Q3").NumberFormat = "0.0%"
   ws.Range("P3").Value = ws.Cells(StartRow, 9).Value
 
 ' ------------------------------------------------------------------------------------------------------------
   GreatesPercentIncrease = 0
   TickerSysmbolGreatesPercentIncrease = " "
   GreatesPercentDecrease = 0
   GreatestTotalVolume = 0
   StartRow = 0
   LastRow = 0
   StartRow = 0
   I = 0
   CheckValues = 0
 
Next ws

End Sub


    
