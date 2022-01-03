# **EXCEL STOCK- ANALYSIS USING VBA** 

##**Overview of Project**

In this project, i am working on refractoring the initial code created using VBA to assist in a Stock analysis. The user wants to be able to analyze more extensive data set and improve the run time of the code.  

The purpose of this project is to refrator the initial script, to enchance its efficiency, and better its run time. The initial code is displayed below

```
4) Loop through tickers   For i = 0 To 11       ticker = tickers(i)       totalVolume = 0       '5) loop through rows in the data       Worksheets("2018").Activate       For j = 2 To RowCount           '5a) Get total volume for current ticker           If Cells(j, 1).Value = ticker Then               totalVolume = totalVolume + Cells(j, 8).Value           End If           '5b) get starting price for current ticker           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then               startingPrice = Cells(j, 6).Value           End If           '5c) get ending price for current ticker           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then               endingPrice = Cells(j, 6).Value           End If       Next j
```

##**Results**

From the anaylsis, the overall stock returns in 2017 were higher than 2018. Only two stocks (Tickers:ENPH and RUN) had a positive return in 2018 while in one Stock (Ticker: TERP) had a negative return.

The initial script had nested loops. In order to make the code more efficient and faster, the nested loops were removed so that the code will loop only once through the dataset while collecting the same information. The refractured code is shown below.

```
'2b) Loop over all the rows in the spreadsheet.For i = 2 To RowCount    '3a) Increase volume for current ticker    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value    '3b) Check if the current row is the first row with the selected tickerIndex.    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value       End If    '3c) check if the current row is the last row with the selected ticker      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value        End If        '3d Increase the tickerIndex.         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then            tickerIndex = tickerIndex + 1         End IfNext i'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.For i = 0 To 11    Worksheets("All Stocks Analysis").Activate    Cells(4 + i, 1).Value = tickers(i)    Cells(4 + i, 2).Value = tickerVolumes(i)    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1      Next i
```
The run time after refractoring decreased from 0.65625 seconds to 0.1328125 seconds in 2017 and 0.67187 seconds to 0.140625 seconds in 2018.

![screenshot of runtime for 2017](https://github.com/kbtwumasi/Stock--Analysis/tree/main/Resources/VBA_Challenge_2017.png)

![screenshot of run time for 2018](https://github.com/kbtwumasi/Stock--Analysis/tree/main/Resources/VBA_Challenge_2018.png)

##**Summary**

Refractoring a code is a way to improve the origianal code. It simplifies a code and can make it efficient, readable, organized and easier to understand. However, refactoring a code can lead to errors being introduced into an already working code. If this happens, it could be time consuming, and waste of money and resources. 

Refactoring the VBA script made the new code more clearer and structured. It made the code efficient as it decreased the run time significantly. The process was however cumbersome and time consuming
