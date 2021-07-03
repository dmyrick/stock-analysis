# Stock Analysis Using VBA

## Overview of Project

### Purpose
The purpose of the Stock Analysis Challenge was to use Excel's Visual Basic for Applications (VBA) to analyze stock data. The analyzation consisted 
of refactoring the VBA script to decrease the time of the script's execution.

## Results

The overall stock performance between 2017 and 2018 decreased in "Total Daily Volume" and "Return" of the stocks. The decrease in Total Daily Volume and Return 
can be seen in the images labeled 2017 and 2018 Stock Performance. 

Code differences which initialized the loops are captured below:

**ticker** 
(green_stocks)

For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0

**tickerIndex**
(VBA_Challenge) 

Dim tickerIndex As Integer
    
        tickerIndex = 0


For i = 0 To 11
        
        tickerVolumes(i) = 0
    
    Next i

Refactoring the VBA script and setting the tickerIndex and tickerVolumes to 0 decreased the time of the script's execution. Ouput arrays were more clearly specified in the. VBA_Challenge.xlsm. 
The differences of the execution times for the original and refactored 2017 and 2018 VBA script runs may be found below. The execution times 
decreased from 2.55 to 0.33 and 2.74 to 0.41, for 2017 and 2018 respectively.

2017 Stock Performance

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Misc_Images/All_Stocks(2017).png)

2018 Stock Performance

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Misc_Images/All_Stocks(2018).png)

2017 Original

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Misc_Images/green_stocks_2017.png)

2017 Refactored

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018 Original

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Misc_Images/green_stocks_2018.png)

2018 Refactored

![ScreenShot](https://github.com/dmyrick/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

1. The advantages of refactoring code are the following:
	- The code is of better quality.
	- The programming is faster.
  
   The disadvantages of refactoring code are the following:
	- The code will have to be retested to ensure the code will run with no syntax errors.
	- There is a risk the programmer may not understand the refactored code, especially if the programmer did not write the original code.

2. Pros of refactoring an original VBA script is the efficiency and speed of receiving the results. In this challenge, the VBA script's run time was decreased. 
   Cons to refactoring an original VBA script is the learning curve to get the end result with no syntax errors.
