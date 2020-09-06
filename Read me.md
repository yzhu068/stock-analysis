# VBA Challenge

## Background: 
Steve, who just graduated with his Finance Degree, is helping his parents to analyze the performace of a stock they are interested: DAQO New Energy CORP. 
In order to diversify the investment, Steve decided to analyze green energy stocks other then DAQO stock. 

### Purpose: 
Steve created an excel which contains the information of 12 stocks' performance in 2017 and 2018. The information including the stock's daily start/close price,daily low/high prices, and total volume.
Steve wants to analyze each stock's performance in regarding to total daily volume and annual return.  

## Analysis and Challenges

### 2017 Performance
 


The analysis of 2017 shows that all the green energy stocks have a very good performance in 2017 except for TERP,which the price dropped by 7.7% on an annual basis. 
DAQO,the stock that Steve's parents interrested, has the best performance in terms of return, which is close to 200%. 
However, DAQO has a relatively low total daily volume comparing to other stocks, indicating it is not yet a popular stock in 2017.

### 2018 Performance
 
Based on the report of 2018 performance, the price of most of the 12 stocks analyzed dropped. DQ in specific, dropped by 62%, which is the highest among the 12 stocks. 

Given that DAQO has the best performance in 2017, it seems this stock is very unstable. So Steve's parents must be cautious when investing on this stock. It is a high risk investment, and if they are looking for short-term return, they may not reach their goal. 

Steve could advise his parents to diversify their investment. Based on two years analyze, ENPH and RUN could be a good choice since they appreciated in both 2017 and 2018. 


### VBA Refactored Code Analysis
After refactoring the code, the runtime of both 2017 and 2018 analysis increased. This can be an indication that the code is less compact and efficient than the original one.



## Summary

-What are the advantages or disadvantages of refactoring code?
Advantages: 
the code can be more readable and understandable to others;
the code may be easier to revise
Disadvantages: 
the code could take longer to execute
errors/bugs may take place while refactoring the code. 


-How do these pros and cons apply to refactoring the original VBA script?
the original VBA is more efficient and give shorter runtime. However, the code is difficult for people to understand other than the developer. The refactor one is more readable, as if the script explains itself, if later Steve wants to expand the analysis to 20, or even more stocks, the refactored code is easier to revise. 
While refactoring the code, I encountered several bugs and took me a while to debug. also the code takes more lines, and chances of making mistakes are higher. 

