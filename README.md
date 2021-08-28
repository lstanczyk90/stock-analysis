# **Green Stock Analysis Project**
## **Overview of the Project**

The purpose of this project was to assist Steve with calcuating the total daily volume and total return for a large group of green stocks. We developed code using VBA to largely automate the process for an existing data set organized by Year, (2017 and 2018) Ticker, Daily Volume, and Price. With the click of a button, Steve may now use our code to quickly calculate the desired outputs in a clean, easy-to-read format. 

While we were able to complete the initial analysis by taking advantage of "nested Ifs" within our code, the resulting calculations were resource-intensive and took a relatively long time to determine the outputs. As such, in Round 2 of our analysis, we did our best to refactor the code by defining an index to utilize in lieu of the nest ifs, resulting in noticeable improvements to our calculation speed.

## **Results**

### **_Original Code_**

    Sub AllStockAnalysis()

        Dim startTime As Single
        Dim endTime As Single


        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer


    'Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate

    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Total Return"

    'Initialize an array of all tickers
    'Define all 12 tickers are string
    Dim tickers(12) As String
        'Assign values to all 12 tickers, starting at 0
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
        
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    
    'Activate the data worksheet.
    Worksheets(yearValue).Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
    'Loop through the tickers
    For i = 0 To 11

        ticker = tickers(i)
        TotalVolume = 0
        
            Worksheets(yearValue).Activate
            For j = 2 To RowCount
                'increase totalVolume if value is ticker
                If Cells(j, 1).Value = ticker Then
                    TotalVolume = TotalVolume + Cells(j, 8).Value
                End If
            
                'Condition to define if previous row is not DQ data row and current row is, then start of DQ data
                'NOTE that in the sheet, it is organized by tickers
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
            
                'Condition to define if next data row is not DQ data row, then ending price
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
                
            Next j
            
        'Output the totalVolume summation onto DQ Analysis
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = TotalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
            
            
            
    Next i
    
    Call FormattingAllStocks

    endTime = Timer
    MsgBox ("This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue))
                    
    
    End Sub


While the above-referenced original code did result in the desired output, we found that the use of nested ifs was detrimental to the performance of the code. Specifically, the 2017 analysis took approximately 0.85 seconds and the 2018 analysis took approximately 0.87 seconds.

### **_Refactored Code_**

    Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers. Use the array to contain all of the 12 tickers, which will allow
    'ticker output and assist in locating starting/ending prices
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index. This index will be used as a point of reference to aggregate outputputs/inputs
    Dim tickerIndex As Single
    tickerIndex = 0
  
    
    

    '1b) Create three output arrays. These will hold the final outputs within All Stocks Analysis
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. Specificaly, the below main loop must reset
    'the tickerVolume to zero when the ticker index is changed to the next value in order to be able to calculate
    'the sum
    For i = 0 To 11
        tickerVolumes(i) = 0

        
        
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    Worksheets(yearValue).Activate
    
    'In order to calculate outputs, code must loop through all rows of the input spreadsheet
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker. Specifically, outputs will be grouped by tickers 0-11.
        'This function will loop until the final row of a ticker is reached, at which point the tickerIndex
        'will shift up by one to create a new output for the new ticker.
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If yes, the value will get stored to the starting price
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
                
        End If
        
        '3c) check if the current row is the last row with the selected ticker.
        'If yes, the value will get stored to the ending price
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            
        End If
        'If the value in the tickers column continues to equal the tickerIndex ticker value, then the code will
        'loop to the beginning of the "For" functionand continue its calculations. If not, it will increase the
        'tickerIndex to introduce the values for the next ticker.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        
        
        
                 
        End If
        
        Next i
        
        'Loop to output stored values into output spreadsheet.
        For i = 0 To 11
        
       
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i
    
    
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

The above-code yielded substantial improvements over the initial method, as evidenced by the below screenshots:

#### **_2017 Output_**

![alt text](https://github.com/lstanczyk90/stock-analysis/blob/3a3de7f2b3dfebe9a27c0dd0efe434f97e05dde3/resources/VBA_Challenge_2017.PNG)

#### **_2018 Output_**

![alt text](https://github.com/lstanczyk90/stock-analysis/blob/3a3de7f2b3dfebe9a27c0dd0efe434f97e05dde3/resources/VBA_Challenge_2018.PNG)

## **Summary**

### **_Advantages and Disadvantages of Refactoring Code_**

The advantages of refactoring code are inherent in the above-referenced results. Refactored code may improve upon the original design, thereby improving on performance and formatting. 

There are, however, disadvantages to consider as well. There is an old addage saying that "if it isn't broken, don't fix it." While refactoring code may yield marked improvements, you are also editing existing code that works. Any time you do this, you introduce potential new bugs, syntax errors, etc. While refactored code may work for an existing data set, it may introduce problems down the road when applied to new use cases. If code is modified over time, you always run the risk of introducing flaws/errors into the code. 

Additionally, the refactored code may not be as intuitive to understand for the average user.

### **_Refactoring the VBA Script_**

As noted above, our refactored code improved the 2017 analysis from 0.85 seconds (original) to 0.17 seconds (refactored), and our 2018 analysis from 0.87 seconds (original) to 0.18 seconds. While it is difficult to appreciate this improvement tangibly (as both outputs took less than one second to generate), consider that the data set we were working with is not that large. For larger data sets, the improvements should be far more noticeable. These improvements are largely the result of our refactored analysis not having to loop over each ticker line-by-line throughout the entire data set. Rather, the refactored code relies on an index to quickly reference and cluster together the appropriate outputs. As such, this such yield great improvements over a large data set.

It can be argued, however, that the refactored code is not as intuitive to the average user as the original. While it makes sense to go go line by line and aggregate volumes and prices by ticker (this is how the average person's mind would visualize the calculation), using an index is not as easy to conceptualize and to follow.




