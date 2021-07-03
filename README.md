# stock-analysis

## Overview of Project
Overview of Project: Explain the purpose of this analysis.
The purpose and background are well defined (2 pt).

In this project, I was tasked to analyze the stock performance of multiple stocks to give Steve and his parents help on how to decide which stocks are best for them. 

## Results
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
The analysis is well described with screenshots and code (4 pt).
![2017 Run Time, Original Code](https://github.com/chloebellehooton/stock-analysis/blob/main/Green_Stocks_2017.png)
![2017 Run Time, Refactored Code](https://github.com/chloebellehooton/stock-analysis/blob/main/VBA_Challenge_2017.png)

The run time was dramatically improved when I refactored the code, as seen in the screenshots for both 2017 and 2018. 
```
    '1a) Create a ticker Index
    'Set to zero before iterating over rows
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
     For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker and adds ticker volume for the current stock ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If it is, then assign current starting price to variable
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
 
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
 
        '3d Increase the tickerIndex if next row's ticker doesn't match previous row's ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1

        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
 ```   

![2018 Run Time, Refactored Code](https://github.com/chloebellehooton/stock-analysis/blob/main/VBA_Challenge_2018.png)
![2018 Run Time, Original Code](https://github.com/chloebellehooton/stock-analysis/blob/main/Green_Stocks_2018.png)

## Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

### What are the advantages or disadvantages of refactoring code?
#### Refactoring code makes the program run faster and smoother. It creates a cleaner and more efficient code. With refactoring, your code takes less steps and doesn't use as much memory. There isn't a clear disadvantage to refactoring other than the opportunity cost of doing so. For some smaller scripts that will only be used a few times, it might not be worth the time to go in and change things for the sake of efficiency. Additionally, if someone new is coming in and refactoring, this could break the code which could be quite costly to the project.

### How do these pros and cons apply to refactoring the original VBA script?
#### Refactoring this was useful because it ran much faster and I ran it for 2 different years so it wasn't a one time script. Also given that I will be sending these to my clients, Steve and his parents, I want to make this code as easy to use as possible. For one, I don't know the memory capabiliities of their computer and we know that Steve is planning to expand the code further so it's best in this case to take the time to refactor it. 
