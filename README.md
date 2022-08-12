# Stock-Analysis
## Goal
The goal here was to take a code that was used to analyze a dataset of only 12 stock market prices. The next step was to refractor the code that instead of being able to analyze just 12 stock market prices it can analyze the prices in the entire stock market in a more efficient way.
### Results 
In order to make the code more efficient I had to change the nesting order for my loops. This was done by creating 4 new arrays
which were tickerVolumes,tickerEndingPrices,tickers, and tickerStartingPrices.

    
    '3) Initialize array of all tickers
    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    '4a) Activate data worksheet
    Worksheets(yearValue).Activate

    '4b) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '5a) Create a ticker Index

    Dim tickerIndex As Single
    tickerIndex = 0

    '5b) Create three output arrays

    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single

    '6a) Initialize ticker volumes to zero
    
     For i = 0 To 11
    tickerVolumes(i) = 0

    Next i
    '6b) loop over all the rows

    For i = 2 To RowCount

    '7a) Increase volume for current ticker
   
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
    
    '7b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
        
        
    End If
    
    '7c) check if the current row is the last row with the selected ticker
    If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
        

        '7d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
        
       End If

     Next i

    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
    Next i
## Summary
### General Code
One main advantage of refactoring code is that the code becomes more efficient and it takes less time to run especially if there is a lot of data to 
analyze and time is limited. One main disadvantage is that it can take a while to test and make sure it is stable and if not refactored correctly there may be bugs.
### Stock Code
The main advantage of refactoring the code from this assignment is that most of the old code is used and very little is added so therefore the code is mostly stable and there are little to no worries about breaking the code. 
One disadvantage is that if there is no strong understanding about how the syntax works, it will be very difficult and time consuming to refactor it because of how 
important the syntax matters when it comes to making it more efficient
