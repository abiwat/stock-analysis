# Refactor VBA Code and Measure Performance

## Overview of Project: Explain the purpose of this analysis.

The client, Steve, is looking to do an analysis of the entire stock market for the last few years to determine which stocks have positive return on investment.  To do this, I am refactoring a macro built to analysis 12 green stocks to report Total Daily Volume and the Return and to expand it to handle the entire stock market.

## Results:

I am going to break the results down to the analysis of the data and then the analysis of the macro.

#### Data

The data made available for this challenge is:
- Ticker - symbol for the individual stock
- Date - the date
- Open - what the stock opened at for that date
- High - the highest value that the stock traded at for that date
- Low -  the lowest value that the stock traded at for that date
- Close - the value that the stock closed at for that date
- Adj. Close - the adjusted value the stock closed at for the date
- Volume - the volume of stock traded for that date

The data that was calculated for this challenge:
- Total Daily Volume - the daily average volume of stock traded
- Return - the net gain or loss of an investment over a specified time

I created three output arrays, tickerVolume(i), tickerStartingPrices(i), and tickerEndingPrices(i) and set the values to 0.  Then created a For loop from i from 0 to 11 to move thourgh the ticker symbols.  
![array.png](https://github.com/abiwat/stock-analysis/blob/main/Resources/Array.png)

I then created a conditional For loop to increase the tickerIndex.  
![Conditional%20Loop.png](https://github.com/abiwat/stock-analysis/blob/main/Resources/Conditional%20Loop.png)

Lastly, I created a For loop to output Ticker, Total Daily Volume and Return.

![Output%20Loop.png](https://github.com/abiwat/stock-analysis/blob/main/Resources/OutPut%20Loop.png)


From the resulting data for 2017, all but TERP would have been a good investment.  That being said, DQ had the highest return (199.4%).  2018 does not appear to be as profitable for the stocks market.  DQ had the lowest return for 2018 (62.6%), ENPH (81.9%) and RUN(84.0%) were the only stocks in the data to have positive return in 2018.  Based on these two years worth of data, I would recommend investing on ENPH and RUN.

#### Macro

The refactoring of macro to increase the amount of data it can handle did make the code run faster on the data available.

The original code for 2017 took 0.4960938 seconds to run, and 2018 took 0.484375 second to run.  With the code refactored, 2017 code ran in 0.1445313 

![2017.png](https://github.com/abiwat/stock-analysis/blob/main/Resources/2017.png)

and 2018 ran in 0.13678188 seconds.  

![2018.png](https://github.com/abiwat/stock-analysis/blob/main/Resources/2018.png)

Clearly the code ran faster after being refactored.

## Summary:

### What are the advantages or disadvantages of refactoring code?

The advantage to refactoring code is the structure of the code and the basic flow is already in place and then you can concentrate on editting and re-writing peices of the code as needed.  It safes time and provides a starting place.  And in refactoring the code, it was cleaned up and ultimately ran faster.

### How do these pros and cons apply to refactoring the original VBA script?

I don't see any cons for this script.  The pro is the amount of time to run the code was reduce.
