# VBA-of-Wall-Street
## Overview of Project ##
The purpose of this project is to utilize VBA in excel to accomplish two general tasks. The first task is to organize and visualize stock ticker values to improve understanding of stock trends between the years of 2017 and 2018. Secondly, by editing the filter code seeing if we can refactor the code such that it improves performance (in this case speeds up the speed of the code)
## Results ##
### Visualization ###
- To enhance visualization of the result, we color coded the annual ticker returns. In this case red shows a year end decrease in returns while green shows a year-end positive ticker return value.
- In addition, we organized the stocks by Ticker and included the daily volume of transactions per ticker to show how much activity each ticker got during the time period.
### Refactoring ###
Refactoring the analysis in order to improve performance.
- Initially, we only utilized the ticker array to compile the code, this meant that every "For" loop had to figure out the ticker volume and ticker prices individually. While this was effective, and reduced initial coding effort, it resulted in a slower run.
``` 
         '5) loop through rows in the data
         Worksheets(yearValue).Activate
         For j = 2 To RowCount
        
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        
            End If
            
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        
            End If
        
            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        
            End If
        
        Next j
```
-Instead of only one array (ticker) we created 3 new ones: volume, starting price, and ending price. By doing this, the VBA code was able to hold the arrays in memory and pull them when needed, improving time considerably.
 ```
    1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
     
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

        '3d) Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
            End If
    
    Next i
```    
By running the new code we went from around 1.85 seconds per analysis down to around 0.277 seconds. this is an 85% improvement. As you can see:
![2017](https://github.com/Z-Zachattack/02-VBA-of-Wall-Street/blob/main/Resources/VBA_Challenge_2017.png)

![2018](https://github.com/Z-Zachattack/02-VBA-of-Wall-Street/blob/main/Resources/VBA_Challenge_2018.png)
## Summary ##
### Advantages of Original ###
The original way of collecting and compiling data had a few interesting advantages, namely: 
- The original VBA code had fewer keystrokes,
- fewer arrays, meaning there were less 'under-the-hood' processes
- Resulting in a more understandable, if only slightly, code. 
### Advantages of Refactoring ###
Refactored VBA code also had a few advantages, such as:
- 85% Faster runtimes
- fewer computer resources allocated
- better equiped to deal with ticker changes/additions

All-in-all, I would say that the refractored code was an improvement over the original. That said, it would have been better to use R Script or Python to collect data so that the tickers wouldn't have to be inputted manually.
