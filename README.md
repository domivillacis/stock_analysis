# Green Stocks Analysis
## Overview of Project
To start this project I have been approached by Steve, who recently got his degree in finance and is helping his parents invest in the stock market. 
His parents are interested in investing solely in green energy stocks (DQ), but Steve thinks his parents should diversify their investments.

In this project, we are going to review the trading performance of the specific stock Steve's parents want to invest in, compared to the performance of all stocks in 2018, and analyse whether or not it is profitable to invest. We are going to use Visual Basic for Applications (VBA) programming language, which will allow us to efficiently code the exact data we need to make this comparison and get relevant data to help Steve's family invest successfully.


## Results: 

To obtain relevant data, we started by creating 3 values that helped us with the calculation: “Year”, “Total Daily Volume” (the total number of shares traded throughout the year) and “Return” (the percentage difference in price from the beginning to the end of the year).

The findings of the first analysis based on the yearly return results is, that DQ trading shares in 2018 dropped by 63%; thus I would advise not investing in it.

<img width="408" alt="DQ_Stocks" src="https://user-images.githubusercontent.com/112814924/194395615-20e29cf6-8dd7-4d71-8073-6c4c60a8ddc1.png">

On the other hand, we have also analyzed the performance of all the stocks in 2018. As we can see, the results of our data are not favorable either, only two stocks had positive tradings, ENPH and RUN. That means there is recesion in the stock market. However, Steve's parents should definitely consider including ENPH and RUN stocks in their investment portfolio.

<img width="670" alt="All_stocks_performance" src="https://user-images.githubusercontent.com/112814924/194395642-c976913b-ed79-4fcc-85c3-eeecf9bc5b6c.png">

### Now let's dive into the technical part of our analysis. 

To obtain all this data, we have made use of different codes such as for loop, variables, conditionals, arrays, among others.

For 2018, we have executed two different codes, which have allowed us to see exactly the same information but at a different running time. In our excel worksheet we have 3013 rows with information, and through the codes used, we have been able to achieve better visibility of the data we need. We must take into account that all this analyses generate space on our computer as well as response time.
Running the first code from year 2018 has taken 0.25 seconds.

<img width="930" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/112814924/194407360-79f4f5d1-1767-4105-9968-d538116ade90.png">

To work more efficiently, we have refactored the code and thus help Steve get the result in the shortest possible time and with the same information.

Through the use of arrays, which are practically lists, and looping through them, we have modified the code a little, to obtain the same data but with a faster running time. In our case it has been 0.07 seconds.

<img width="929" alt="VBA_Challenge_R_2018" src="https://user-images.githubusercontent.com/112814924/194407419-cf207271-5d73-40a1-b17f-57a17ee58bcc.png">


## Summary:

As a final result, we can see that we have executed the same analysis with different codes, both have worked but one has been more efficient than the other.

One of the advantages of refactoring a code is that we can simplify it to improve the result of data collection or running times; however, we can fall into a serious confusion if we do not do it correctly.

Changing a code also means creating confusion for the next programmer who has access to this information. For this reason, it is always a good idea to write notes to point out how the code is built. Making use of whitespace and indentation can also help make the code readable and understandable.

Our current Excel worksheet contains a significant amount of information but at the same time this worksheet does not contain as much information as other worksheets could. Therefore, code simplification can definitely be helpful in finding critical data and improving running times.
