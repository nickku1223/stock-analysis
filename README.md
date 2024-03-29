# Stocks Analysis with VBA

## Project Overview

The purpose of this project is to use VBA to create marco so the stakeholders can easily get the results for the stocks analysis of the year they desired.


## Results

![2017 Stocks Analysis Result](resources/2017_Stocks_Analysis_results.png)

From the results of the 2017 stocks analysis, we can see all of the stocks but "TERP" has positive returns, and the top 3 performers are 

- "DQ" with 199.4% of return, 
- "SEDG" with 184.5% of return and 
- "ENPH" with 129.5% of return. 

Overall, 2017 is a good year for most of the stocks we have listed.

![2018 Stocks Analysis Result](resources/2018_Stocks_Analysis_results.png)

With the results we get from the 2018 stocks analysis, we can see most of the stocks have the negative returns, the top 3 of the worst peformers are

- "DQ" with -62.6% of return, 
- "JKS" with -60.5% of return, 
- "SPWR" with -44.6% of return. 

But we can also see **"ENPH"** still has **81.9%** of return, during 2017 - 2018 **"ENPH"** has a total **211.4%** of return which made it the top performer out of all stocks. The other stock that has a positive return for both years is **"RUN"**, in 2017, it has a small 5.5% of return, but in 2018 it gained momentum and has the highest return of **84%**.

### Runtime for original script and refactored script
<sub>2017 Original Script Runtime</sub>
![2017 Original Script Runtime](resources/2017_AllStocksAnalysis_Original_Script_runtime.png)
<sub>2018 Original Script Runtime</sub>
![2018 Original Script Runtime](resources/2018_AllStocksAnalysis_Original_Script_runtime.png)
<sub>2017 Refactored Script Runtime</sub>
![2017 Refactored Script Runtime](resources/VBA_Challenge_2017.png)
<sub>2018 Refactored Script Runtime</sub>
![2018 Refactored Script Runtime](resources/VBA_Challenge_2018.png)

## Summary

With the result we got from both 2017 and 2018, we can see that "ENPH" and "RUN" would be good choices to invest in for the long term during the 2 years, the top performers such as "DQ", "SEDG" could be good for a short term but have to closely monitor the performance and set a target price and/or stop loss point.

#### 1. What are the advantages or disadvantages of refactoring code?

The advantage is that it reduced the runtime. The disadvantage for me personally is that at first I was confused about the code by adding tickerIndex and using it to access the tickers array, and why to create the 3 out put as an array, but I think creating output as arrays helps to process and store the respective array elements, which helps with the runtime.

#### 2. How do these pros and cons apply to refactoring the original VBA script?

The big difference between the two scripts is that setting output arrays, and using arrays have some advantages, such as it can reduce the runtime, all the array elements being stored in the continuous memory location, and the code could be easy to maintain. (some of the benefits are the results of a google search,and [this is the webpage](https://vbaf1.com/tutorial/arrays/advantages-of-an-array/#:~:text=Array%20Advantages%3A&text=All%20the%20array%20elements%20are,using%20array%20index%20or%20subscript.)

Cons on the other hand, for me is that the code can seem rather complicated and less straightforward at first, but also forced me to look into the code and figure out the relationship between elements, but also learned and explored a different, and more efficient way to complete the same work.
