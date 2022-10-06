# Green Stocks Analysis
## Overview of Project
For this data project, I've been approached by Steve, whom recently got his degree in finance and is helping his parents invest in the stock market. His parents are interested in investing solely in green energy stocks, but Steve thinks his parents should diversify their investment.

In this project we are going to review the trading performance of different stocks in 2018 and 2018, and analyse whether it is profitable to invest or not. By using Visual Basic for Applications (VBA) programming language, we will be able to effitiently code the exact data we need to make this comparison and we will obtain relevant data to help Steve's family invest successfully.


## Results: 

For 2018, I started analysing specifically the shares traded by the DQ stock. To obtain this data, 3 values that helped us with the calculation were created: “Year”, “Total Daily Volume” (the total number of shares traded throughout the year) and “Return” (the percentage difference in price from the beginning to the end of the year).

As a result, the findings of the first analysis is that DQ shares in 2018 fell by 63%, and that means, right now is not the best time to invest in this green stock. However, 2 other stocks performed better in that.

<img width="408" alt="DQ_Stocks" src="https://user-images.githubusercontent.com/112814924/194395615-20e29cf6-8dd7-4d71-8073-6c4c60a8ddc1.png">

On the other hand, I have also analyzed the performance of all the stocks in 2018. As we can see, the data of our results are not favorable, only two stocks that are ENPH and RUN had positive traded shares.

<img width="670" alt="All_stocks_performance" src="https://user-images.githubusercontent.com/112814924/194395642-c976913b-ed79-4fcc-85c3-eeecf9bc5b6c.png">

We are now going to dive into the technical part of our analysis. To obtain all this data, we have made use of different codes such as for loop, variables, conditionals, arrays, among others.

For 2018, we have executed two different codes, which have allowed us to see exactly the same information but at a different execution time. In our excel worksheet we have 3013 rows with information, and through the codes used, we have been able to achieve better visibility of the data we need. However, all this analysis generates space in our computer and response time.
Running the the first code from year 2018 has taken 0.25 seconds.



And to work more efficiently, we have refactored the code and thus help Steve to obtain the calculation in the shortest possible time and with the same information.

Through the use of arrays, which are practically lists, and looping through them, we have changed the code a little, to obtain the same data but that the running time is the best possible, in our case it has been 0.07 seconds.



## Summary:

As a final result, we can see that we have executed the same analysis with different codes, both have worked but one has been more efficient than the other.

One of the advantages of refactoring a code is that we can simplify it to improve the result of data collection or running times; however, we can fall into a serious confusion if we do not do it correctly.

Changing a code also means creating confusion for the next programmer who has access to this information. For this reason, it is always a good idea to write notes to point out what exactly that written code is working on and making use of whitespace and indentation can definitely help make code readable and understandable.

In our current Excel spreadsheet we have important information to get the required data but at the same time this worksheet does not contain as much information as other spreadsheets. So code simplification can definitely be helpful in getting data and improving running times.
