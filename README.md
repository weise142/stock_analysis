#Stock Analysis using VBA

##Overview of Project

1. Create a VBA macro to read, change, format cells, and create timer pop up messages to understand how long it took to run our macro
2. Use for loops and nested for loops to reduce code run times
3. Apply coding skills to refactor our macro code to reduce the run time, which is an important part of coding. This is especially important when working with large datasets

##Purpose

The purpose of the assignement is to assist Steve in an analysis of the performance of Green Energy Stocks in 2017 and 2018. The original code was written with for loops that looped through multiple times to determine the total daily volume and return percentage of 12 Green Energy Stocks. We then refactored the code to determine if the changes to the for loops to loop through only once would reduce the run time.

In our Excel spreadsheet, we created run buttons that would populate an input box so Steve could choose which year he would like to run an analysis on. We also coded a timer that provided feedback on how quickly the code ran and populated a message box with this information once the macro ran completely. This exercise helped familiarize us with for loops, nested for loops, arrays, if-then conditional statements, assigning data types, and formatting. 

#Results

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/2017%20Stock%20Performance.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/2018%20Stock%20Performance.png)
Left: Stock Performance 2017, Right: Stock Performance 2018

In comparing the returns for these stocks between 2017 and 2018, there are a few factors we need to look at. The total daily volume(TDV) of these stocks did not have an impact on the returns. For example, SEDG and JKS had similar TDV in 2017 and 2018 but the return percentages were 53.9% and -60.5% for JKS and 184.5% and -7.8% for SEDG in 2017 and 2018 respectively. TDV was also not relevant to how how much a high or low a stocks returns were respectively. The highest return in 2017 was DQ at 199.4% despite having the lowest TDV, while RUN had the best return in 2018 of 84% despite having the 3rd highest TDV. Although the TDV does not appear to have an impact on the return percentage for a specific year, it does help illustrate potential volatility of a stocks based on its outstanding share float. Generally stocks such as DQ and HASI are going to be more volatile than stocks such as FSLR and SPWR because the delta value per share of the stock is going to be lower due to a larger outstanding share float. 

This analysis is helpful as well, because it helps illustrate the vlatility and total returns of these Green Energy Stocks over 2017 and 2018. The top performing stocks were ENPH, SEDG, and DQ, so these stocks could pose good investments for Steve's parents as they invest in Green Energy Stocks. 

##Run Time Comparison: Original vs Refactored Code

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Initial%202017%20run%20time.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Initial%202018%20run%20time.png)
Left: Run Time of original code 2017, Right: Run Time of original code 2018

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Refactored%202017%20run%20time.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Refactored%202018%20run%20time.png)
Left: Run Time of refactored code 2017, Right: Run Time of refactored code 2018


