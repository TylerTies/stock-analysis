# Stock Analysis

## Overview of Project:
The purpose of this analysis was to refactor the code in the existing stock analysis file to determine if we can increase the efficiency.  The goal was to speed up the code to allow for analysis of more than just a handful of stocks.  Refactoring allowed us to collect the same information but in a more efficient manner.

## Results
As far as stock performance is concerned, 2017 fared significantly better than 2018.  Only one of the selected stocks, TERP, saw negative return in 2017 of -7.2%.  On the other hand, in 2018 there were only two stocks that had a positive return, ENPH and RUN.  The rest saw losses, some significant.

Some refactoring of the code led to significant improvements in performance.  Processing times for the code dropped to around 0.1 seconds compared to just under a full second with the previous code

![2017](/Resources/VBA_Challenge_2017.png)
![2018](/Resources/VBA_Challenge_2018.png)
### Previous Code
Refactoring the code allowed us to eliminate some redundancy in working through the rows of data.  The previous code had a For loop to cycle through the tickers with another For loop nested inside to work through every row of data.  This meant each row was being touched once for each ticker.  This redundancy might be insignificant for smaller data sets but quickly can become burdensome as your rows increase.
### Refactored Code
The refactored code introduces a tickerIndex that allowed us to store individual values for each ticker as we go through the code and saved for later when we create the output.  An example is below where the tickerIndex is used to count and store the tickerVolumes.

```
If Cells(i, 1).Value = tickers(tickerIndex) Then
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
End If
```

This allowed us to remove one of the For loops and require only one cycle through the code.  The output values were stored in arrays instead of individual variables that had previously needed to be output before being reused again in another For loop running through all rows.  The refactored code ran significantly faster.

## Summary
Refactoring the code had the advantage of starting with a solid foundation.  It was code that had already been used successfully but just needed some tweaking and rearranging.  This was much easier than starting from scratch. Refactoring the existing code also saved time by not having to look up solutions that have already been solved. One disadvantage could be if the code was confusing, messy, or lacked comments to help understand what is happening.  In this case it might be very time consuming just to understand the current code and see if it's worth refactoring.  You could spend hours of effort just to understand and add your own comments to the old code.  In this analysis there weren't many cons that I experienced by choosing to refactor the original script.

The original script had a solid foundation but was inefficient.  It had the advantage of being slightly less confusing but lacked the ability to scale in size efficiently.  The refactored script added some arrays and made the code slightly more difficult to understand, but this was solved by ensuring appropriate comments were included.  The new script allows for scalability and performs more efficiently




