
# An Analysis of Wall Street stock market 

## Overview of Project

This project automates an analysis on an Excel dataset stock market using VBA 

### Purpose
To help Steve look into green energy stocks including DQ for his parents to invest 



## Analysis and Results

The analysis is based on total daily volume of each stock and their return rate for 2017 and 2018.

### Analysis of stocks in 2017

The highest daily volume in 2017 belonged to SPWR and the lowest to DQ.

DQ had the highest return rate (199.4%) and TERP was the only one with negative return rate(-7.2%)

In 2017 , 11 stocks out of 12 had a positive return rate.  

![Click Here to see the results for 2017]()

### Analysis of stocks in 2018

2018 was a challenging year for most of green energy stock market compared to 2017. Out of 12, only two stocks had positive return rates.

The highest daily volume belonged to ENPH and the lowest was AY.

The highest return rate in 2018 also beloned to ENPH at 81.9%. DQ had a negative return rate at -62.6% despite a 200% growth in daily volume.  

![Click Here to see the results for 2018](resources/) 


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

- Disadvantages: The process time was slower because of nested loops and keep activating different worksheets within the loops. In this code we could not access and reuse the data in memory that was created in the loop for each ticker.   


**The factored code:**

- Advantages: Shorter process time. Used an array of 12 to keep the total volume and start price and end price for each ticker. That would give us an access to reuse the data stored in the array.It didn't use a nested loop and didn't jump into different worksheet while generating data.

- Disadvantages: It used a more complicated data structure to keep the data. It was not easy and fast to code and took time to design our variables.



 #### Lida Ghalam
 #### December 2020