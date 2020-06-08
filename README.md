# Excel VBA of Wall Street
### Author: Lucy Wo

![stock Market](pictures/stockmarket.jpg)

### Background

Use VBA scripting to analyze real stock market data to find insights.

### File

From 2014 to 2016 stock data, including ticker, date, high, low, close and volume.

### Steps

Create a script that will loop through all the stocks for one year for each run and take the following informations:

1. Unique ticker symbol.
2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.
5. Conditional formatting that will highlight positive change in green and negative change in red.

### Challenges

* Be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"
* Allow it to run on every worksheet, i.e., every year, just by running the VBA script once

### Conclusions

* From 2014 to 2016, number of unique tiker has increased to 3168
* The changes of greatest % increase from 2016 versus 2014 elevated 6093%
* More positive change in 2016
* Overall 2016's stock performance better than 2015 and 2014

### Scripts of VBA 
<pre>
Sub summarefort()

Dim Ws As Worksheet 
For Each Ws In Worksheets 

Dim i, j, Lastrow As Long 
Dim yearchange, percentchange, maxpercent As Double
Dim total As LongLong
Dim openp, closep As Double

Ws.Range("O2").Value = "Greatest % increase"
Ws.Range("O3").Value = "Greatest % Decrease"
Ws.Range("O4").Value = "Greatest total volume"
Ws.Range("P1").Value = "Ticker"
Ws.Range("Q1").Value = "Value"
Ws.Range("I1").Value = "Ticker"
Ws.Range("J1").Value = "Yearly Change"
Ws.Range("K1").Value = "Percent Change"
Ws.Range("L1").Value = "Total Stock Volume"

j = 2
total = 0
Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
openp = Ws.Cells(2, 3).Value

For i = 2 To Lastrow 
    If Ws.Cells(i, 1).Value <> Ws.Cells((i + 1), 1).Value Then
        Ws.Cells(j, 9).Value = Ws.Cells(i, 1).Value
        Ws.Cells(j, 12).Value = total + Ws.Cells(i, 7).Value
        closep = Ws.Cells(i, 6).Value
        yearchange = closep - openp
        Ws.Cells(j, 10).Value = yearchange
        If yearchange > 0 Then
            Ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            Ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        If openp <> 0 Then 
                    percentchange = (yearchange / openp) * 100
        End If 
        Ws.Cells(j, 11).Value = percentchange 
        Ws.Cells(j, 11).Value = (CStr(percentchange) & "%")
        total = 0 
        j = j + 1 
        openp = Ws.Cells(i + 1, 3).Value
        percentchange = 0
    Else
        total = total + Ws.Cells(i, 7).Value
    End If     
Next i

'------------part two-------------------------------------

Dim astsu As Integer
Dim k As Long
Dim matchin, matchde As Integer 
Dim ginp, gdep As Double  
Dim o As Long
Dim maxtotal As LongLong
astsu = Ws.Cells(Rows.Count, "K").End(xlUp).Row 
ginp = 0 
gdep = 0
maxtotal = 0 

For k = 2 To astsu 
    If Ws.Cells(k, 11).Value >= ginp Then
        ginp = Ws.Cells(k, 11).Value
        Ws.Cells(2, 17).Value = ginp 
         Ws.Cells(2, 16).Value = Ws.Cells(k, 9).Value
         Ws.Cells(2, 17).Value = (CStr(Ws.Cells(2, 17).Value * 100) & "%")
    End If 
    If Ws.Cells(k, 11).Value <= gdep Then
        gdep = Ws.Cells(k, 11).Value 
        Ws.Cells(3, 17).Value = gdep 
        Ws.Cells(3, 16).Value = Ws.Cells(k, 9).Value 
        Ws.Cells(3, 17).Value = (CStr(Ws.Cells(3, 17).Value * 100) & "%") 
    End If 
    If Ws.Cells(k, 12).Value >= maxtotal Then 
        maxtotal = Ws.Cells(k, 12).Value 
        Ws.Cells(4, 17).Value = maxtotal 
        Ws.Cells(4, 16).Value = Ws.Cells(k, 9).Value 
    End If
    
Next k 

Ws.Columns("A:Q").AutoFit

Next Ws
End Sub
<pre>
