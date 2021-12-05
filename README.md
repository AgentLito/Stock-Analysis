# Stock-Analysis

Stock Analysis with Excel VBA

Overview of Project

Purpose

This task was to analyze 12 stocks from 2017 and 2018 and determine their total daily volume and yearly return for each stock. Looking at the total daily volume we can get a snapshot of how actively a particular stock is being traded. The yearly return will allow us to measure how well the stock performed from beginning of the year to the end of the year.

Results
Stock Performances between 2017 and 2018

A Look at 2017 Total Daily Volume
When we analyze the stock for 2017, the results provide a variety of stocks that performed well with varying degree of total daily volume and return in the positive direction.
All Stocks (2017)		
		
Ticker	Total Daily Volume	Return
AY	136,070,900	8.9%
CSIQ	310,592,800	33.1%
DQ	35,796,200	199.4%
ENPH	221,772,100	129.5%
FSLR	684,181,400	101.3%
HASI	80,949,300	25.8%
JKS	191,632,200	53.9%
RUN	267,681,300	5.5%
SEDG	206,885,200	184.5%
SPWR	782,187,000	23.1%
TERP	139,402,800	-7.2%
VSLR	109,487,900	50.0%

If 2017 data is all that we had in total, then our analysis would give us a few stocks that could possibly be investment candidates but with only the one year sample. It leaves the decision to invest a more difficult one to process. Thankfully we also have 2018 data to compare and analyze to 2017, so let's look at how these stocks did in 2018 to see how we would have done.

A Look at 2018 Total Daily Volume
In 2018 the story  seems to have a completely different one.
All Stocks (2018)		
		
Ticker	Total Daily Volume	Return
AY	83,079,900	-7.3%
CSIQ	200,879,900	-16.3%
DQ	107,873,900	-62.6%
ENPH	607,473,500	81.9%
FSLR	478,113,900	-39.7%
HASI	104,340,600	-20.7%
JKS	158,309,000	-60.5%
RUN	502,757,100	84.0%
SEDG	237,212,300	-7.8%
SPWR	538,024,300	-44.6%
TERP	151,434,700	-5.0%
VSLR	136,539,100	-3.5%

We see that there are only two stocks that performed better overall for the year – RUN and ENPH. But let's now try to compare the two years and see what the data tells us.
Lets get a glimpse in comparing the two data sets we have combined the tables and have two additional columns with some calculations based of what we know. By adding the column Difference in Total Daily Volume, where we take the volumes for 2017 and compared to the volumes of 2018, it is clear that there are seven stocks that more activity in 2018 than 2017. But to get an even clearer picture of what stocks did the best over the two years lets sort the Return Over 2 Years with the best results at the top. The stock ENPH had the best return over two years with 211.45% followed by SEDG, DQ and RUN.

Summary

What are the advantages or disadvantages of refactoring code?
The purposes of refactoring according to Martin Fowler (Father of Code Smell) are stated in the following:
1.	Refactoring Improves the Design of Software
2.	Refactoring Makes Software Easier to Understand
3.	Refactoring Helps Finding Bugs
4.	Refactoring Helps Programming Faster

It also allows the code to be more adaptive over time and allow for the ability to bring new developers into the code without much training if refactored correctly.
The disadvantage of refactoring code is that it takes money and time which may not be available or limited.
How do these pros and cons apply to refactoring the original VBA script?

The disadvantages just mentioned don’t come into play in our exercise, but all of the advantages did. Let's address them one at a time.
1.	Refactoring Improves the Design of Software in our exercise allowed us to make the code more streamline and allowed for the script to address if the data set were to change or grow and the run time is greatly improved thus not taxing the memory of the system and making the user experience much better.
2.	Refactoring Makes Software Easier to Understand in our exercise by reducing the amount of code and adding more concise commenting on the script, it allows for us to revisit this code later down the road or someone entirely new to review the code and have a good understanding of what is trying to be achieved.
3.	Refactoring Helps Finding Bugs which in our exercise there were a few that cropped up but because the code was paired down, it allowed us to find the issues and resolve them quickly.
4.	Refactoring Helps Programming Faster because now that the code is streamlined and commented, it allows for us or anyone else to add functional code to improve and expand the capabilities of the script.

Original Code
Sub AllStocksAnalysis():

    Dim startTime As Single
    Dim endTime  As Single

yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer

 '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   
   '3b) Activate data worksheet
   Sheets(yearValue).Activate
   
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Sheets(yearValue).Activate
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
       
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

  Call formatAllStocksAnalysisTable
  
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Sub formatAllStocksAnalysisTable():

'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00"
   Range("C4:C15").NumberFormat = "0.00%"
   Columns("B").AutoFit
   
   dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
    
End Sub

Refactored Code

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
    'Initialize array of all tickers
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
    'Activate data worksheet
    Worksheets(yearValue).Activate
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    '1a) Create a ticker Index
    tickerIndex = 0
    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrice(11) As Single
    Dim tickerEndingPrice(11) As Single
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrice(tickerIndex) / tickerStartingPrice(tickerIndex) - 1
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

