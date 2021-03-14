### Stock-Analysis

#Overview of the Project

A stock is an investment that can have a great value to the investors. Therefor, in order to help the investors to make the right purchase we are here to analyse the dataset given to us. 

We have dataset of daily volume exchanged per stock, for the year of 2017, and 2018 which can easily reflect on value of the stock. Dataset included over 3000 daily records that can help us to to serve interest of client named Steve.

#Results

First of all, our client Steve needed particular stock called "DADQ" with the ticker 'DQ'.
Here we created code for the total yearly traded volume based on the daily volumes for the ticker 'DQ'.
 
  Sub DQAnalysis()
Worksheets("DQAnalysis").Activate
Range("A1").Value = "DAQO (Ticker: DQ)"
'Create a header row
Cells(3, 1).Value = "Year"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"  
Worksheets("2018").Activate
totalVolume = 0

Dim startingPrice As Double
Dim endingPrice As Double
rowStart = 2
'DELETE: rowEnd = 3013
'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

For i = rowStart To rowEnd
   'increase totalVolume
    If Cells(i, 1).Value = "DQ" Then
       totalVolume = totalVolume + Cells(i, 8).Value
       End If
    If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
           startingPrice = Cells(i, 6).Value
       End If
       If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
           endingPrice = Cells(i, 6).Value
       End If
Next i
   Worksheets("DQAnalysis").Activate
   Cells(4, 1).Value = 2018
   Cells(4, 2).Value = totalVolume
   Cells(4, 3).Value = (endingPrice / startingPrice) - 1
End Sub




