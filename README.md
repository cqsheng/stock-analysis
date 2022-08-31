# An Analysis of stock performances

## Overview of Project
    This project is used to determine the different performances of various stock performances for a given year, the total stock traded volume and the return rate for the stock.
### Purpose
    The purpose is to determine if there are any stocks that would be good investments based on their return rates for their past performances. As well as pros and cons of refactoring the VBA code to reduce the runtime of the script used to calculate the stock performances.

## Results
    
### Stock performance comparison
    Although the exact numbers changed, the same two stocks, ENPH and RUN produced a positive return for both 2017 and 2018 while every other stock produced negative returns.
    
![](/images/VBA_Challenge_2017_stocks.png)
![](/images/VBA_Challenge_2018_stocks.png)
### Refactored script execution times
    As can be seen on the pictures below, the code refactoring helped improve the performance of the script drastically, the runtime for the code before(shown on the left) takes more than 5 times to run than the code after its refactored(shown on the right).
    This is done by removing the nested for loops which loops over the entire dataset for each ticker, into one for loop which loops over the data just once, and using different arrays to store the value for all the different stocks.
    
![](/images/VBA_Challenge_2018.png)
![](/images/VBA_Challenge_2018_Refactored.png)
## Summary
    
### 1.What are the advantages and disadvantages of refactoring code?
    The general advantages to refactoring the code is that it's done to make the code more efficient, by taking less steps for the the code to run, using less memory or less steps. The disadvantage is that refactoring takes more time on the developers side, even though there is already a working version of the code, you'll be spending more time in making a more efficient refactored code, which could be much more difficult to write from a logical standpoint than the original code that was first made to solve the problem.
### 2.How do these pros and cons apply to refactoring the original VBA script?
    To use this challenge as an example, the pros for the refactored code is that it runs significantly faster, more than 5 times faster on the same dataset on my computer. The cons is that it takes longer to actually code for, requiring the creation of new variables and arrays. As well as knowing the the original code and data well enough to be able to create a version of the code that doesn't require the nested for loops to analyze the results in one pass of the data.


