# An Analysis of stock performances

## Overview of Project
    This project is used to determine the different performances of various stock performances for a given year, the total stock traded volume and the return rate for the stock.
### Purpose
    The purpose is to determine if there are any stocks that would be good investments based on their return rates for their past performances. As well as pros and cons of refactoring the VBA code to reduce the runtime of the script used to calculate the stock performances.

## Results
    
### Stock performance comparison
   2017 was a much better stock performance year than 2018, there was only one stock with a negative return rate in 2017 and that was only by -7.2%, while all the other 11 stocks had positive returns. By contrast 2018 only had 2 stocks with positive return rates, resulting in a much worse performing year overall.
    
### Refactored script execution times
    As can be seen on the pictures below, the code refactoring helped improve the performance of the script drastically, the runtime for the code in 2017(shown on the left) takes slightly more time to run than for 2018(shown on the right). However both were much faster than the non refactored code, taking less than a fifth of the time to run compared to them.
    This is done by removing the nested for loops which loops over the entire dataset for each ticker, into one for loop which loops over the data just once, and using different arrays to store the value for all the different stocks.
    
![](/images/VBA_Challenge_2017.png)
![](/images/VBA_Challenge_2018.png)
## Summary
    
### 1.What are the advantages and disadvantages of refactoring code?
    The general advantages to refactoring the code is that it's done to make the code more efficient, by taking less steps for the the code to run, using less memory or less steps. The disadvantage is that refactoring takes more time on the developers side, even though there is already a working version of the code, you'll be spending more time in making a more efficient refactored code, which could be much more difficult to write from a logical standpoint than the original code that was first made to solve the problem.
### 2.How do these pros and cons apply to refactoring the original VBA script?
    To use this challenge as an example, the pros for the refactored code is that it runs significantly faster, more than 5 times faster on the same dataset on my computer. The cons is that it takes longer to actually code for, requiring the creation of new variables and arrays. As well as knowing the the original code and data well enough to be able to create a version of the code that doesn't require the nested for loops to analyze the results in one pass of the data.


