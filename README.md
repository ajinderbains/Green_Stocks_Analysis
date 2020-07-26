# Green_Stocks_Analysis

## Purpose :

This report will help Steve to analyze the stock market for past years to make decision to invest in future. Report included total daily volumes and return for each share in past years.
Additional this report also analyzed the advantages and disadvantages of using nested and single loop.

## Results: 

Excel sheet is used ,which has data for stocks for years  2017 and 2018.This excel sheet also include shares ,daily opening and closing prices along with daily Volume.
To perform the analysis, total volume of each share is calculated .Starting and ending price is used for each share to calculate return for each share.

## 2017 Stock Returns

![2017 Stocks Return](https://github.com/ajinderbains/Green_Stocks_Analysis/blob/master/Resource/Stocks2017.PNG)

## 2018 Stock Returns


![2018 stocks return](https://github.com/ajinderbains/Green_Stocks_Analysis/blob/master/Resource/Stocks2018.PNG)




-	In 2017 stock market was better than 

## Summary: 


### Refactored code:
 #### Advantage
1.	Only one loop is used and it reduces the complexity. It takes less time and is efficient way when calculations and writing output to excel sheet is done separately. Because of one loop for all calculations for total volumes and starting and ending prices for each ticker take noticeable less run time.

2.	Values for total daily volumes ,starting and ending are calculated and  stored so they can be used for another analysis  like to find percentage each ticker/stock of total daily volumes in comparison to total volumes of stock market .

 

#### Disadvantages:
1.	Memory is used in storing  values for total daily volumes, starting and ending price .So execution time  is less though but there is memory cost.

### Old Code/Nested loop:
#### Advantages:
1.	As output is written directly  to excel sheet so there is no memory cost/used to store values. Fast runtime is traded for memory.

#### Disadvantages:
1.	With nested loops execution time is longer as compare to one loop only.
2.	For new programmers it is difficult to understand nested loops.
3.	When calculations and output is written on excel sheet at same time then it increases the runtime.


