# stock-analysis
_UMN VBA homework_

## Purpose
The client is helping his parents examine stock data from 2017 and 2018 to determine which companies to invest in. The client's parents believe the following:
>. . .if a stock is traded often, then the price will accurately reflect the value of the stock.

So I want to create a macro that examines the total daily volume, which is the total number of shares traded in a day, and the yearly return, or the price difference between the beginning and ending of the year, for each company in each year to help the client examine the stock data and present the results to his parents.

## DQ Analysis:
His parents were first interested in a company called DQ. So, I made a macro to analyze the 2018 sheet for their performance. This way I could check that my code is able to get the total daily volume and yearly return into a simple table of a single stock for a single year.

1. I created a sub called DQAnalysis which activated a worksheet called "DQ Analysis” and created a title at the top of the worksheet.
```
Sub DQAnalysis()    
    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker: DQ)"
````
2. Then I created a header row for the data and activated the worksheet with the stock data for 2018.
```
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate
```
3. I initialized the total volume at 0 by creating a variable called totalVolume and setting it equal to 0. 
```
    totalVolume = 0
```
4. I declared two variables as doubles for the starting and ending prices. The starting price is the closing price on the first day of the year and the ending price is the return on the last day of the year which will be used to give the yearly return as an output.
```
    Dim startingPrice As Double
    Dim endingPrice As Double
```
5. I created variables to determine what rows the for loop should loop through.
```
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
```
6. I created a for loop that loops through the data in 2018. If the company equals "DQ" it increases the total volume by the value for that day. It establishes the starting price of that stock by determining that if the current row equals "DQ" and the previous one does not. Similarly, the ending price of that stock can be determined if the current row equals "DQ" and the next one does not.
```
     For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        
        'Calculate starting Price
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If

        'Calculate ending Price
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If
```
7.  Then I activated the "DQ Analysis" worksheet under the header I printed the year, total volume (that was calculated in the for loop), and the yearly return which is calculated by dividing the ending price by the starting price and then subtracting one. Then I ended the sub.
```
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
End Sub
```
##### Results of DQ Analysis:
![DQ Analysis Table](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/DQ_Analysis.png)

 It turned out DQ's 2018 performance wasn't very good it was traded 107,873,900 with a yearly return of -63%, but I was able to ensure my for loop is able to collect the total volume and the yearly return for a single stock in a single year. Now I can use it to make a table to summarize the performance of all the companies for each year. This will make it easier for our client to add new data each year and compare the results of multiple companies.

## All Stocks Analysis:
1. First I created a new macro called AllStockAnalysis. In it I created a variable called YearValue that asks for an input value for the year this will be used to determine which data the analysis calls. I also defined some variables that i will use to see how fast this code runs.
```
Sub AllStocksAnalysis()
   YearValue = InputBox("What year would you like to run the analysis for?")
   Dim startTime As Single
   Dim endTime As Single
    startTime = Timer
```    
2.  Then I created a title, header row, and did some formatting.
```
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + YearValue + ")" 
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"   
    'Format Title
    Cells(1, 1).Font.Name = "Ariel"
    Cells(1, 1).Font.Bold = True
    Cells(1, 1).Font.Size = 14    
    'Formatting header
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.Size = 12
    Range("A3:C3").Font.Name = "Ariel"
    Range("A3:C3").Font.Color = RGB(255, 255, 255)
```
3. I created an array that will go through the 12 companies whose performance I'm analyzing.
```
   Dim tickers(11) As String
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
```

4. I declared two variables as doubles for the starting and ending prices as I had done in the DQ Analysis. 
```
    Dim startingPrice As Double
    Dim endingPrice As Double
```
5. I  activating the worksheet this time substituting "2018" with the YearValue variable so the end user can select the year. I used the same variables as before to determine what rows the for loop should loop through.
```
   Worksheets(YearValue).Activate
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
```
6. Then I created a for loop to run through the tickers and initialize the totalVolume to zero.
```
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       Worksheets(YearValue).Activate
```
7. I made a nested for loop to run through the rows. This is basically the same for loop as the one I created in the DQ Analysis sub. The only difference is really that I am using the variable ticker instead of "DQ." 

```
       For j = rowStart To RowCount
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
      Next j
```
 8. The data output was the same as for the DQ Analysis, but the active sheet was the All Stock Analysis instead. Also the first column now contains the company names instead of the year. Then I ended the outer loop for the tickers.
```
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   Next i
```
 9. I added some formatting to make the data look a little prettier and make it easier to read. This included making the columns fit to size, changing some number formatting and added some conditional formatting.
```
        'Formatting Table Text
    Range("A4:C15").Font.Name = "Ariel"
    Range("A4:C15").Font.Size = 10
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    
    'table formatting
    'https://software-solutions-online.com/excel-vba-cell-fill-color/
    Range("A3:C3").Interior.Color = RGB(0, 0, 0)
    Range("A15:C15").Borders(xlEdgeBottom).LineStyle = xlContinuous

    Columns("A").AutoFit
    Columns("B").AutoFit
    Columns("C").AutoFit
   
    'Conditional formatting
    dataRowStart = 4
    DataRowEnd = 15
    
    For i = dataRowStart To DataRowEnd
        If Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = RGB(100, 0, 0)
            Cells(i, 3).Font.Color = RGB(255, 0, 0)
        ElseIf Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = RGB(0, 100, 0)
            Cells(i, 3).Font.Color = RGB(0, 255, 0)
        End If
    Next i
```
10. Before ending the sub I finished the timer and created a message box to display the data.
```    
    endTime = Timer
    MsgBox "The " & YearValue & " code ran in " & (endTime - startTime)
End Sub
```
11. I checked that my code worked and then I created a sub to clear the data from the worksheet. Then I added two buttons to the All Stocks Analysis Spreadsheet. One to run the AllStocksAnalysis macro and one to run the ClearWorksheet macro. 
```
Sub ClearWorksheet()
    Cells.Clear
End Sub
```

## Results: All Stocks Analysis
![All Stocks Analysis 2017 Table](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/All_Stocks_Analysis_2017.png)

![All Stocks Analysis 2018 Table](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/All_Stocks_Analysis_2018.png)

From just looking at the reports I've prepared for the client we can deduce the following in general there was a higher trade volume in 2018 than 2017, but the yearly returns tended to be lower. This seems counterintuitive to the general consensus that higher trade volume means higher sustainability. However, in the best performers RUN and ENPH this construct remained true. They were both heavily traded and had positive returns for both 2017 and 2018. If we are just going off of this data those would be the stocks I would recommend. It is 2022 so there is the basic issue of this data being a little out of date, but I think the client should also consider expanding the report to include the following for the last 5 to 10 years to exclude outliers that could be caused by something simple like a high year starting price:
- The starting price and ending price.
- The average price for the year.
- choose a period instead of year start and year end. Examine the stock through that period.

_Reference:_
Butler, R. A. (2022, March 10). How to evaluate stock performance. Investopedia. Retrieved April 10, 2022, from https://www.investopedia.com/articles/investing/011416/how-evaluate-stock-performance.asp 

## All Stocks Analysis Refactor

![All Stocks Analysis Timer Message Box](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/Unrefactored_Timer_Results.png)

Our first AllStockAnalysis macro ran a little slow. To see if I could get the code to run a little faster, I did a refactor.

##### Pros to Refactoring Code:
- Makes code faster.
- Makes code easy to read.
- Makes code easier to add additional function to later.
- Helps to keep code up to date.
- Fixing bugs in the original code.

##### Cons to Refactoring Code:
- Refactoring code takes time and time is money.
- There is a chance you will make a Mistake. In this example I am using VBA so I can see easily if the code is broken, but in another language or on a bigger project it may be more difficult to debug the code if you break it while refactoring.
- May be difficult for beginners.

1. The beginning is the same as the AllStocksAnalysis macro:
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    Worksheets("All Stocks Analysis").Activate   
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'create header
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

    Dim tickers(11) As String
    
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
    
    Worksheets(yearValue).Activate
    
    StartRow = 2
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
 ```   
2. First I created a tickerIndex and sit it to zero this will be a reference for the ticker array and the output arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices which I defined below.
```
    tickerIndex = 0
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
```    
 3. I created a for loop to initialize the three arrays to zero.
```
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
```
 4. Then I used the framework of the nested for loop from the AllStocksAnalysis to create a for loop that uses the tickerIndex to create arrays for the Ticker Volume, Ticker Starting Price and Ending Price. 
```
   ' Loop over all the rows in the spreadsheet.
    For i = StartRow To RowCount
	'increase ticker volume
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
	   'determine ticker closing price for ticker on first day of the year.
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           End If     
	   'determine ticker closing price for ticker on last day of the year.
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
           End If
```
5. Then I added an if statement to move to go to the next ticker once it looped through all the rows for the current ticker.
```
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerIndex = tickerIndex + 1
            End If                    
    Next i
```    
6. I created a loop based on the AllStocksAnalysis macro and subbed in my new variable names to output the arrays' data to a table. I also formatted the table and added the message box for the timer.
```
    For i = 0 To 11        
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1       
    Next i

    Worksheets("All Stocks Analysis").Activate
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
```
## All Stocks Analysis Refactor: Results

![All Stocks Analysis Refactored 2017 Table](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/All_Stocks_Analysis_Refactored_2017.png)

![All Stocks Analysis Refactored 2018 Table](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/All_Stocks_Analysis_Refactored_2018.png)

##### Pros to the Refactored VBA Code
I was able to get my data to match the original AllStocksAnalysis macro and the refactor was faster. 
- The refactored code ran faster because in the original code the nested for loop is switching back and forth between worksheets to gather and store data, but the refactored code is able to use the tickerIndex to store the data into arrays. Once the data is gathered it can use the arrays to populate the data onto the other sheet, so it doesn't need to switch back and forth in between each ticker. If you are working with a large data set this refactor will significantly benefit you. 
- Another benefit of the tickerIndex variable is that it removes the need for nested loops which makes the code easier to read. One disadvantage of the original code was the use of nested loops. Nested loops and if statements can be difficult to read and understand how the code works. So using the tickerIndex improves the readability of the code by assigning consistent easy to read variable and getting rid of the nested loop.
- With VBA you can test the refactor easily.

![[All Stocks Analysis Refactored Timer Message Box](https://github.com/MichelaZ/stock-analysis/blob/main/Submission/Refactored_Timer_Results.png)

##### Pros for the Original Code:
- Although the original code was much slower than the refactor, if you were working with a small data set I don't think it's worth taking the additional time to refactor the code. The functionality is the same and with even with this size of a data set the original code took less than a second. I don't think the average person would really notice much difference between .15 seconds and .78 seconds.
- Developing the original code made it a lot easier to develop the refactored code.

##### Cons for the Refactored Code:
- Took time to refactor.
- Might be difficult for beginners.

##### Cons for the Original Code:
- The code was slower.
- The code was difficult to read.

_Author's notes:_  
- All the files I created are in the submission folder. The resources folder just contains the downloads.
- The formatting is a little different between the AllStocksAnalysis and the AllStocksAnalysisRefactored macro, but the reason I started my code over from the starter code was to get it to look more like the assignment prompt. 
- The actual macro has comments inside the code. I removed most of the ones in my excerpts, because I added further explanation in the numbered steps, so I found them to be redundant. If you would like to see my comments in the code, please open the TXT file or the macro enabled workbook. 
- Another step I might take to improve both the refactored and the original code would be to improve the code for the yearValue variable. If you don’t enter one of the sheet names it might give the user error as it is now. 
