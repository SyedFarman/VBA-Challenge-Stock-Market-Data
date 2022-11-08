# VBA-Challenge-Stock-Market-Data

## Overview of Project

The purpose of this project is to analyse daily Stock Market Data, in which we are anlysing top performing tickers, percentage change and total stock volume for the year 2018 to 2020
 
## Results

In this project, i have summrised the ticker by yearly change in price, perctange change and total volume and the stock which performed well or worsed during the year and the highest stock volume


***1.a)‘Creating Loop Through all worksheets to automate to process in every work sheet.***
```
Dim ws As Worksheet
For Each ws In Worksheets
```
***b)‘Counting Last Row.***
```
RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
```
  

***2.a)'Set an initial variable for holding the Tricker name.***
```
Dim Ticker_Name As String
 ```
  ***b)'Set an initial variable for holding the total Stock per Tricker.***
   ```
Dim Total As Double
Dim TotalStock As Double
 ```

***c)‘Setting the Total to zero.***
 ```
TotalStock = 0
 ```

***3)‘Keep track of the location for each Tricker & it's balances in the summary table.***
 ```
  ws.Cells(1, 10).Value = "Ticker"
  ws.Cells(1, 11).Value = "Yearly Change"
  ws.Cells(1, 12).Value = "Percent Change"
  ws.Cells(1, 13).Value = "Total Stock Volume"
      
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
 ```

 ***4.a)‘Loop through all stocks.***
 ```
  For i = 2 To RowCount
 
 ```
 ***b)‘Check if the current row is the first row with the selected ticker and assign the current starting price to the OpenPrices variable.***
 ```
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
OpenPrice = Cells(i, 3).Value
End If
```
***c)'check if the current row is the last row with the selected ticker and assign the current closing price to the ClosePrice Prices variable.***
```
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ClosePrice = Cells(i, 6).Value
End If
 ```
 ***d)'Set the Tricker.***
 ```
Ticker_Name = Cells(i, 1).Value
```
 ***e)'Set the  Yearly Change.***
 ```
 'YearlyChange = ClosePrice - OpenPrice
```
 ***f)'Set the  TotalStock.***
 ```
 TotalStock = TotalStock + (Cells(i, 7).Value)
```
 ***g)'Set the  Percentage Change.***
 ```
 PercentChange = FormatPercent(YearlyChange / OpenPrice)
```

***5.)'Print the TickerName, YearlyChange,PercentChange and TotalStockVolume in the Summary Table .***
```
 ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
 ws.Range("K" & Summary_Table_Row).Value = YearlyChange
 ws.Range("L" & Summary_Table_Row).Value = PercentChange
 ws.Range("M" & Summary_Table_Row).Value = TotalStock
 ```

***6.)'Conditional Formatting positive change in green and negative change in red .***
```
 If Range("K" & Summary_Table_Row).Value <= 0 Then
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
Else
ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
End If
```

***7.a)'Finding Greatest Increase,Decrease in Stock, also Greatest Volume .***
```
 Max = Application.WorksheetFunction.Max(Range("L:L").Value)
 Min = Application.WorksheetFunction.Min(Range("L:L").Value)
 TSV = Application.WorksheetFunction.Max(Range("M:M").Value)
 ws.Cells(2, 17).Value = FormatPercent(Max)
 ws.Cells(3, 17).Value = FormatPercent(Min)
 ws.Cells(4, 17).Value = TSV
```
 ***b)‘Finding Tricker Name Greatest Increase,Decrease in Stock, also Greatest Volume.***
 ```
ws.Range("P2").Value = Application.WorksheetFunction.Lookup(ws.Range("Q2").Value, ws.Range("L2:L3001"), ws.Range("J2:J3001"))
ws.Range("P3").Value = Application.WorksheetFunction.Lookup(ws.Range("Q3").Value, ws.Range("L2:L3001"), ws.Range("J2:J3001"))
ws.Range("P4").Value = Application.WorksheetFunction.Lookup(Range("Q4").Value, ws.Range("L2:L3001"), ws.Range("J2:J3001"))
```


## Analysis

As a result of the above codes, We have analysed three years of stock market data. There are three thousand tickets traded in stock market. I have higlighted the good performing trickers in green color and bad performers are in red color, we can also anaylze the performance by using the yearly percentage change and total Volume column. The screenshots attached below for the reference.



##### 2018 Stock Analysis 

![](resources/StockSummary - 2018.png)


I have also analysed the highest performing & worst performing stock and volume of highest performing

The screenshots attached below for the reference.


##### 2018 Highest & Worst performing stock

![](resources/GreatestStock - 2018.png)




