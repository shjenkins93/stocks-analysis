# VBA Challenge Stocks Analysis
# Overview of Project
# Purpose
The purpose of this project was to refractor the previously written VBA script to make it run faster when analyzing stock data from 2017 and 2018. The background of the product was to help Steve analyze stock data from 12 stocks during the years 2017 and 2018 to see if the stocks are worth investing in.
# Results
The results concluded that the code we edited to improve the speed of our macro was successful. Below are the results from the original module code.

<img width="267" alt="Screen Shot 2022-10-20 at 2 33 20 PM" src="https://user-images.githubusercontent.com/113859036/197030384-eead5469-b806-43ea-abcd-9eb788d2e3ca.png">
<img width="271" alt="Screen Shot 2022-10-20 at 2 33 03 PM" src="https://user-images.githubusercontent.com/113859036/197030403-24005a58-367a-424f-aa2d-5e8fb35cfc33.png">
With the edits made, my VBA code now looks like this:

tickerIndex = 0

Dim tickerVolumes(12) As Long

Dim tickerStartingPrices(12) As Single

Dim tickerEndingPrices(12) As Single

For i = 0 To 11

    tickerVolumes(i) = 0
    
    tickerStartingPrices(i) = 0
    
    tickerEndingPrices(i) = 0
    
Next i

For i = 2 To RowCount

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
     
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     End If
     
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerIndex = tickerIndex + 1
            
        End If
        
Next i

For i = 0 To 11  

    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    
    Cells(4 + i, 2).Value = tickerVolumes(i)
    
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1    
    
Next i  

        Worksheets("All Stocks Analysis").Activate
This edited code did increase the speed of the macro as shown below:

<img width="262" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/113859036/197031308-1dbe1670-fa72-4def-a062-7726feff9911.png">

<img width="267" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/113859036/197031345-7b059044-1a45-42b0-b65d-5041dfa735a3.png">

# Summary
Refactoring does have its pros and cons. The pros are that the code is now more concise and not as long. It is easier for the data analyst to follow the code and be able to make changes. In the event of a bug, which I did encounter a few times while building the code, the refactored code made it easier to identify where the bug was and what was stopping the macro from running through the whole code. A con of refactoring the code in a macro is it might not have the power to run large data sets with a lot of variables. 
In this case, refactoring our VBA script improved the speed of the macro because there was less script for the macro to process.




