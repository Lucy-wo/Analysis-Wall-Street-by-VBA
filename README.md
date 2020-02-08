# VBAStocks
###### Here is my scripts for VBA homework
<pre>

Sub summarefort()

Dim Ws As Worksheet <br />
For Each Ws In Worksheets <br />
 <br />
Dim i, j, Lastrow As Long <br />
Dim yearchange, percentchange, maxpercent As Double <br />
Dim total As LongLong <br />
Dim openp, closep As Double <br />
 <br />
Ws.Range("O2").Value = "Greatest % increase" <br />
Ws.Range("O3").Value = "Greatest % Decrease" <br />
Ws.Range("O4").Value = "Greatest total volume" <br />
Ws.Range("P1").Value = "Ticker" <br />
Ws.Range("Q1").Value = "Value" <br />
Ws.Range("I1").Value = "Ticker" <br />
Ws.Range("J1").Value = "Yearly Change" <br />
Ws.Range("K1").Value = "Percent Change" <br />
Ws.Range("L1").Value = "Total Stock Volume" <br />
 <br />
j = 2 <br />
total = 0 <br />
Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row <br />
openp = Ws.Cells(2, 3).Value <br />
 <br />
For i = 2 To Lastrow <br />
    If Ws.Cells(i, 1).Value <> Ws.Cells((i + 1), 1).Value Then <br />
        Ws.Cells(j, 9).Value = Ws.Cells(i, 1).Value <br />
        Ws.Cells(j, 12).Value = total + Ws.Cells(i, 7).Value <br />
        closep = Ws.Cells(i, 6).Value <br />
        yearchange = closep - openp <br />
        Ws.Cells(j, 10).Value = yearchange <br />
        If yearchange > 0 Then <br />
            Ws.Cells(j, 10).Interior.ColorIndex = 4 <br />
        Else <br />
            Ws.Cells(j, 10).Interior.ColorIndex = 3 <br />
        End If <br />
        If openp <> 0 Then <br />
                    percentchange = (yearchange / openp) * 100 <br />
         End If <br />
        Ws.Cells(j, 11).Value = percentchange <br />
        Ws.Cells(j, 11).Value = (CStr(percentchange) & "%") <br />
        total = 0 <br />
        j = j + 1 <br />
        openp = Ws.Cells(i + 1, 3).Value <br />
        percentchange = 0 <br />
         <br />
    Else <br />
        total = total + Ws.Cells(i, 7).Value <br />
    End If <br />
Next i <br />
 <br />
'---part two----- <br />
 <br />
Dim astsu As Integer <br />
Dim k As Long <br />
Dim matchin, matchde As Integer <br />
Dim ginp, gdep As Double <br /> 
Dim o As Long <br />
Dim maxtotal As LongLong <br />
astsu = Ws.Cells(Rows.Count, "K").End(xlUp).Row <br />
ginp = 0 <br />
gdep = 0 <br />
maxtotal = 0 <br />
 <pre>
For k = 2 To astsu <br /> 
    If Ws.Cells(k, 11).Value >= ginp Then <br />
        ginp = Ws.Cells(k, 11).Value <br />
        Ws.Cells(2, 17).Value = ginp <br />
         Ws.Cells(2, 16).Value = Ws.Cells(k, 9).Value <br />
         Ws.Cells(2, 17).Value = (CStr(Ws.Cells(2, 17).Value * 100) & "%") <br />
    End If <br />
  <pre>
    If Ws.Cells(k, 11).Value <= gdep Then <br />
        gdep = Ws.Cells(k, 11).Value <br />
        Ws.Cells(3, 17).Value = gdep <br />
        Ws.Cells(3, 16).Value = Ws.Cells(k, 9).Value <br />
        Ws.Cells(3, 17).Value = (CStr(Ws.Cells(3, 17).Value * 100) & "%") <br />
    End If <br />
     <br />
    If Ws.Cells(k, 12).Value >= maxtotal Then <br />
        maxtotal = Ws.Cells(k, 12).Value <br />
        Ws.Cells(4, 17).Value = maxtotal <br />
        Ws.Cells(4, 16).Value = Ws.Cells(k, 9).Value <br />
    End If <br />
Next k <br />
 <br />
Ws.Columns("A:Q").AutoFit <br />
 <br />
Next Ws <br />
End Sub <br />
<pre>
