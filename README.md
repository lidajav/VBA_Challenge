
# An Analysis of Wall Street Stock Market 

## Overview of Project

This project automates an analysis on an Excel dataset stock market using VBA 

### Purpose
To help Steve look into green energy stocks performance including DQ stock for his parents to invest 



## Analysis and Results

The analysis is based on total daily volume of each stock and their return rate for 2017 and 2018.

### Analysis of stocks in 2017

The highest daily volume in 2017 belonged to SPWR and the lowest to DQ.

DQ had the highest return rate (199.4%) and TERP was the only one with negative return rate(-7.2%)

In 2017 , 11 stocks out of 12 had a positive return rate.  

[Click Here to see the results for 2017](Resources/VBA_Challenge_2017.png)

### Analysis of stocks in 2018

2018 was a challenging year for most of green energy stock market compared to 2017. Out of 12, only two stocks had positive return rates.

The highest daily volume belonged to ENPH and the lowest was AY.

The highest return rate in 2018 also belonged to ENPH at 81.9%. DQ had a negative return rate at -62.6% despite a 200% growth in daily volume.  

[Click Here to see the results for 2018](Resources/VBA_Challenge_2018.png) 


## Results

- The return rate for DQ stocks had a sharp decline in 2018 compared to 2017 which would make it very risky to invest.

- ENPH looks like a promising stock to invest. The return rates have been positive in both years and was one of the only two stocks with a positive rate ( The highest) in 2018 


## Summary

### Advantages and disadvantages of refacoring code in general
  
**Advantages:** 

- Shorter process time and more efficient

- Takes fewer steps and less memory

- Easier to read and reuse

**Disadvanages :**

- May use more complicated data structure

- Take more time and effort to write a code with improved logic

 ### Advantages and disadvantages of the original and refactored VBS script
 
 **The original code:** 

- Advantages: The original code was easier to write. It used only 3 variables for all tickers to hold total volume and startprice and endprices. 

- Disadvantages: The process time was longer because of nested loops (36,156 loops versus 3013 loops) and keep activating different worksheets within the loops. The Total Daily Volume and Starting Price and Ending Price for each ticker will not be saved and reused. 


**The factored code:**

- Advantages: Shorter process time. Used three arrays of 12 to keep the Total Daily Volume, Starting Price and Ending Price for each ticker and for that reason there was no need for the nested loop. Therefore there was no need to keep switching worksheets while generating data. 

- Disadvantages: It used a more complicated data structure to keep the data.We used 36 variables ( 3 arrays of 12) instead of 3 variables.  It was not easy and fast to code and took time to design our variables and output.



 #### Lida Ghalam
 #### December 2020
