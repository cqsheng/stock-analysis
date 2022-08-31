# An Analysis of stock performances

## Overview of Project
    This project is used to determine the different performances of various stock performances for a given year, the total stock traded volume and the return rate for the stock.
### Purpose
    The purpose is to determine if there are any stocks that would be good investments based on their return rates for their past performances. As well as pros and cons of refactoring the VBA code to reduce the runtime of the script used to calculate the stock performances.

## Results
    
### Stock performance comparison
    Although the exact numbers changed, the same two stocks produced a positive return for both 2017 and 2018 while every other stock produced negative returns.
    
![](/images/VBA_Challenge_2017_stocks.png)
![](/images/VBA_Challenge_2018_stocks.png)
### Refactored script execution times
    As can be seen on the pictures below, the code refactoring helped improve the performance of the script drastically, the runtime for the code before(shown on the left) takes more than 5 times to run than the code after its refactored(shown on the right).
    This is done by removing the nested for loops which loops over the entire dataset for each ticker, into one for loop which loops over the data just once, and using different arrays to store the value for all the different stocks.
    
![](/images/VBA_Challenge_2018.png)
![](/images/VBA_Challenge_2018_Refactored.png)
## Summary
    
### 1.What are the advantages and disadvantages of refactoring code?
    
### 2.How do these pros and cons apply to refactoring the original VBA script?



