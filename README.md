# Green_Stocks_Analysis

New  code:

Advantage
1.	Only one loop is used and it reduces the complexity. It takes less time and is efficient way when calculations and writing output to excel sheet is done separately. Because of one loop for all calculations for total volumes and starting and ending prices for each ticker take noticeable less run time.

2.	Values for total daily volumes ,starting and ending are calculated and  stored so they can be used for another analysis  like to find percentage each ticker/stock of total daily volumes in comparison to total volumes of stock market .

 

Disadvantages:
1.	Memory is used in storing  values for total daily volumes, starting and ending price .So execution time  is less though but there is memory cost.

Old Code:

Advantages:
1.	As output is written directly  to excel sheet so there is no memory cost/used to store values. Fast runtime is traded for memory.

Disadvantages:
1.	With nested loops execution time is longer as compare to one loop only.
2.	For new programmers it is difficult to understand nested loops.
3.	When calculations and output is written on excel sheet at same time then it increases the runtime.
