# Stock-Analysis

Performing analysis on green stocks values

## Project Overview

To help Steve in analyzing the green stocks for his parents to invest in and to check if it is worth investing in. To get through the analyzing, I used Visual Basic Appication in Excel to find the 
stock's total daily volumes and annual returns.First, I found the details about only one stock and then I compared other 11 more green stocks to see what would be the best option for his parents.

Compairing all the stocks gave a detailed study which could be understood by Steve's parents.

### Purpose

The purpose of this project is to make an analysis, which will show up the multiple stocks in an effective way using VBA. To find out the details needed in an effective way which also includes refactorization of the code, which makes the analysis efficient.

So, this project shows up whether the refactoring can make the analysis more efficient

## Result

### Refactoring the code

To refactor the code, I created a variable tickerIndex and used with four different arrays; tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices. With the help of loops and conditional loops, the output arrays are getting their values. Instead of looping through the sheet values again and again as in the code without refactoring, here after refactoring the code, it is made simple and easy to understand.


#### Original Code

Sub AllStocksAnalysis()

    'Initializing the startTime and endTime
    Dim startTime As Single
    Dim endTime  As Single

   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   
   'Creating a variable yearValue and taking Input from user
   yearValue = InputBox("What year would you like to run the analysis on?")
    
     startTime = Timer
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   
  'Range("A1").Value = "All Stocks (2018)"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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

   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
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
   '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       
       Cells(4 + i, 2).Value = totalVolume
       
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
          
        
   Next i
   

In the original code the control shifts twice between the worksheets and also because of the loop structure, it goes through the whole worksheet again and again and hence it takes more time to calculate the values/results.

#### Refactored Code

Sub AllStocksAnalysisRefactored()

    'Initializing the startTime and endTime
    Dim startTime As Single
    Dim endTime  As Single
    
    'Creating a variable yearValue and taking Input from user
    yearValue = InputBox("What year would you like to run the (refactored code)analysis on?")

    'Initializing the variable startTimer to the current time
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All StocksRF (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes, tickerStartingPrices, tickerEndingPrices to Zero
    
        For tickerIndex = 0 To 11
            tickerVolumes(tickerIndex) = 0
            tickerStartingPrices(tickerIndex) = 0
            tickerEndingPrices(tickerIndex) = 0
        Next tickerIndex
        
        'Reinitialize tickerIndex to zero
        tickerIndex = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
             
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                    
        '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
                    
            End If
            
        'End If
    
        Next i
    
        Worksheets("All Stocks Analysis").Activate

   
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    

The variable allowed me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to 
each ticker symbol before iterating through the data set. After applying this code, the analysis is
completed much faster as it loops through all the rows and simaltaneously takes the volumes for each ticker and calculates the output.

### Run-time for both methods and years

The links to the screenshots are given below which shows the working of the original code.

https://github.com/Vaishali715/Stock-Analysis/blob/main/Resources/2017_Original.png

https://github.com/Vaishali715/Stock-Analysis/blob/main/Resources/2018_Original.png

The links to the screenshots are given below which shows the working of the refactored code.

https://github.com/Vaishali715/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png

https://github.com/Vaishali715/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png

Looking at the run-time values it is understood that the refactored code is efficient than the original code.

## Summary

### Advantages and Disadvantages of Refactoring code in general

The **advantages** of refactored code is that it helps to make the code clean and organized so it is easy to understand  and it may happen that a some bugs may appear which are then easy to fix while refactoring.

The **disadvantage** of refactored code is that we are changing the code that already works and there could be a chance that we may not refactor it correctly, in that case it is a waste of time and efforts. It is risky when the application is big.


### Advantages and Disadvantages of Original and Refactored code in VBA script

The **advantage** of original and refactored VBA script is that we can use the original code fully to develop our clean and organized refactored code. It is easy to understand and works faster than the original code.

The **disadvantage** is that if the code is big then it becomes difficult with respect to time and money to be spent on it.


Stock_3 testing 456
