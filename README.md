# VBA-of-Wall-Street
## Overview of Project ##
The purpose of this project is to utilize VBA in excel to accomplish two general tasks. The first task is to organize and visualize stock ticker values to improve understanding of stock trends between the years of 2017 and 2018. Secondly, by editing the filter code seeing if we can refactor the code such that it improves performance (in this case speeds up the speed of the code)
## Results ##
### Visualization ###
- To enhance visualization of the result, we color coded the annual ticker returns. In this case red shows a year end decrease in returns while green shows a year-end positive ticker return value.
- In addition, we organized the stocks by Ticker and included the daily volume of transactions per ticker to show how much activity each ticker got during the time period.
### Refactoring ###
-We refactored the analysis in order to improve performance.
- Initially, we only utilized the ticker array to compile the code, this meant that every "For" loop had to figure out the ticker volume and ticker prices individually. While this was effective, and reduced initial coding effort, it resulted in a slower run.
>       '5) loop through rows in the data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
     >      '5a) Get total volume for current ticker
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
The analysis is well described with screenshots and code (4 pt).
## Summary ##
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
