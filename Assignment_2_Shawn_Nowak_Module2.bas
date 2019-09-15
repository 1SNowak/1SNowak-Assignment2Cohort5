Attribute VB_Name = "Module2"
Sub StockMarketAnalysis()

    ' Set an initial variable for holding the Ticker
    Dim Ticker As String
  
    ' Set an initial variable for holding the total per stock
    Dim Stock As Double
    Stock = 0
    
    ' Add Headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
    ' Keep track of the location for each stock in Ticker column I
    Dim Ticker_Column As Integer
    Ticker_Column = 2
    
    ' Loop through all stocks for the year
    For I = 2 To 760192
    
        ' Check if we are still within the same stock, if not...
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
        ' Set the Stock name/ticker
        Ticker = Cells(I, 1).Value

        ' Add to the Stock
        Stock = Stock + Cells(I, 7).Value

        ' Print the Ticker symbol in the ticker column I
        Range("I" & Ticker_Column).Value = Ticker
        
        ' Print the Total Stock Volume to the column J
        Range("J" & Ticker_Column).Value = Stock
        
        ' Add one to the Ticker_Column row
        Ticker_Column = Ticker_Column + 1

            ' Reset the Total Stock
            Stock = 0
            
        ' If the cell immediately following a row is the same Stock...
        
        Else
        
        ' Add to the Stock
        Stock = Stock + Cells(I, 7).Value

        End If

    Next I
    
End Sub


