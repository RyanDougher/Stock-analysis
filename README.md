# An Analysis of Green Stock
## Overview of Project

### Purpose

In this project, I used VBA to creat an analysis of stock market data for 12 different "green" stocks. The analysis finds the total daily volume and annual return for each stock, over the years of 2017 and 2018. I then refactored the code to perform the same analysis more efficiently, and making it easier for future users to read.

## Results

### Stock Performance

Comparing the overall stock performances of 2017 and 2018, we can see nearly all stock performed better in 2017. TERP and RUN were the only two stocks that did better in 2018. While DQ did well in 2017, it dropped by 62.6% in 2018. ENPH and RUN were the only two that showed positive returns for both years, with ENPH beating RUN as the best performing stock over the course of both years.

! [2017 Stock Performance](/blob/main/Resources/Stock_Performance_2017.png)

! [2018 Stock Performance](/blob/main/Resources/Stock_Performance_2018.png)

### Refactored Code

While the original code worked to run the analysis, it can be improved on to run more efficiently. I was able to make the code more efficient by switching the nesting order of the loops. I created arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. I established "tickerIndex" as a new variable and used this to match the arrays with the tickers array. As shown in tbe images below, this reduced the run time of the code from about 0.6 seconds to 0.1 seconds, which is an improvement on the original code.

These images show the run time of the orignal code:

![2017 Original Code Run Time](/blob/main/Resources/Resources/VBA_Challenge_2017.png)

![2018 Original Code Run Time](/blob/main/Resources/Resources/VBA_Challenge_2018.png)

These images show the run time of the refactored code:

![2017 Refactored Code Run Time](/blob/main/Resources/Resources/VBA_Challenge_2017.png)

![2018 Refactored Code Run Time](/blob/main/Resources/Resources/VBA_Challenge_2018.png)

##Summary

### Advantages and Disadvantages of Refactoring

The main advantage of refactoring the code is that it runs more efficiently, and is easier to read. This helps if we ever need to use the same code on a larger data set, or if someone in the future needs to understand the code. The disadvantage is that restructuring a code that is already working is time consuming, with a small risk that it does not run effectively. In order to avoid that, it is smart to save the original code in a place where it can be easily accessed later on.

### Refactoring in VBA

Specifically in VBA, a refactored code results in a shorter run time. This is an advantage, especially when working with larger data sets. Also, the refactored code looks more organized and is easier for a reader to digest.
