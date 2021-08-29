# stock-analysis
VBA Module 2.1


## Overview of Project
### Purpose
The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information from the years 2017 and 2018 in order to determine whether or not the stocks are worth investing into. Our goal was to improve the efficiency of the code we practiced in the module.

---

## Results
### Analysis
First, I copied the code from the given README file. Then, I completed the required steps and filled in the missing parts of the code, which have been copied below:
 
 '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays 
    Dim tickerVolumes(12) As Long 
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
 #### Run-Time Pop-Ups and Charts
 ![VBA_Challenge_2017](/Users/irinapreotescu/Desktop/Resources/VBA_Challenge_2017.png)
 ![VBA Chart 2017](/Users/irinapreotescu/Desktop/Resources/VBA Chart 2017.png) 
 
 ![VBA_Challenge_2018](/Users/irinapreotescu/Desktop/Resources/VBA_Challenge_2018.png)
 ![VBA Chart 2018](/Users/irinapreotescu/Desktop/Resources/VBA Chart 2018.png) 

---

## Summary
### Advatnages and Disdvantages of Refactoring Code
An advantage of clean code is faster and more efficient programming, design, debugging, and improved software. It also facilitates cooperation by making the code easier to read by others. 

A disadvantage is that applications that are too large might not have test cases due to volume, which might make refactoring code more difficult and can actually be counterproductive.

### Advantages and Disadvantages of the Original and Refactored VBA Script
Advantages include an improved run-time of the macro (from 5-8 seconds to less than 1 second), a debugging of the code, and clearer instructions.
The original code had several disadvantages that refactoring fixed, speeding up the analysis.
