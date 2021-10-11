# stocks-analysis
Performing an analysis on stock info for more informed decision making

## Overview
The purpose of this project was to help Steve analyse the dataset of 2017 and 2018 stock data.
I refactored the code for faster execution and cleanliness. I aldo added buttons to clear the cells
and run analyses for both years separately. With the new refactored code, Steve can now expand his search
to much more than the original dataset of 12 stocks- it is now possibl for him to run the analysis on 
thousands of stocks at a time.

## Results
The resulting refactored code cleaner, better organized, and executed faster then the original, as shown:

The results also show that the initial hypothesis that DQ is a great pick is faulty. The stock demonstrated
more volatility than many of the others and did not demonstrate positive trends or traits that set it apart.
Steve will have to run this analysis on a wider range of stocks with the refactored code- If this sample is
representative of the other companies in it's industry, then many other better options should be revealed.

## Summary
The refactoring of code in general poses several advantages: It is easier to read and edit in the future,
it uses clearer variables whose purposes are evident by their names, and it eliminates waste in terms of
unnecessary functions. 
When looking specifically at the code in this analysis, I created variables for the current, previous, and
next row's stock tickers (currentRowsTicker, prevRowsTicker, nextRowsTicker) I did this so I didn't have to
repeatedly use the expressions below over and over to make the code easier to navigate and reference:
Cells(Row, 1).Value (current row's value) 
Cells(Row - 1, 1).Value (previous row's value)
Cells(Row + 1, 1).Value (next row's value)
This is one of many examples of how the refactored code doesn't just execute faster, but can be edited and
read more efficiently as well.

