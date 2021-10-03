#Stock Analysis using VBA

##Overview of Project

1. Create a VBA macro to read, change, format cells, and create timer pop up messages to understand how long it took to run our macro
2. Use for loops and nested for loops to reduce code run times
3. Apply coding skills to refactor our macro code to reduce the run time, which is an important part of coding. This is especially important when working with large datasets

##Purpose

The purpose of the assignement is to assist Steve in an analysis of the performance of Green Energy Stocks in 2017 and 2018. The original code was written with for loops that looped through multiple times to determine the total daily volume and return percentage of 12 Green Energy Stocks. We then refactored the code to determine if the changes to the for loops to loop through only once would reduce the run time.

In our Excel spreadsheet, we created run buttons that would populate an input box so Steve could choose which year he would like to run an analysis on. We also coded a timer that provided feedback on how quickly the code ran and populated a message box with this information once the macro ran completely. This exercise helped familiarize us with for loops, nested for loops, arrays, if-then conditional statements, assigning data types, and formatting. 

#Results

##Green Stocks Analysis

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/2017%20Stock%20Performance.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/2018%20Stock%20Performance.png)
Left: Stock Performance 2017, Right: Stock Performance 2018

In comparing the returns for these stocks between 2017 and 2018, there are a few factors we need to look at. The total daily volume(TDV) of these stocks did not have an impact on the returns. For example, SEDG and JKS had similar TDV in 2017 and 2018 but the return percentages were 53.9% and -60.5% for JKS and 184.5% and -7.8% for SEDG in 2017 and 2018 respectively. TDV was also not relevant to how how much a high or low a stocks returns were respectively. The highest return in 2017 was DQ at 199.4% despite having the lowest TDV, while RUN had the best return in 2018 of 84% despite having the 3rd highest TDV. Although the TDV does not appear to have an impact on the return percentage for a specific year, it does help illustrate potential volatility of a stocks based on its outstanding share float. Generally stocks such as DQ and HASI are going to be more volatile than stocks such as FSLR and SPWR because the delta value per share of the stock is going to be lower due to a larger outstanding share float. 

This analysis is helpful as well, because it helps illustrate the vlatility and total returns of these Green Energy Stocks over 2017 and 2018. The top performing stocks were ENPH, SEDG, and DQ, so these stocks could pose good investments for Steve's parents as they invest in Green Energy Stocks. 

##Run Time Comparison: Original vs Refactored Code

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Initial%202017%20run%20time.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Initial%202018%20run%20time.png)
Top: Run Time of original code 2017, Bottom: Run Time of original code 2018

![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Refactored%202017%20run%20time.png)![This is an image](https://github.com/weise142/stock_analysis/blob/9e6fd17c16a46779f7a9e674a376ca1bd4f573d3/Refactored%202018%20run%20time.png)
Top: Run Time of refactored code 2017, Bottom: Run Time of refactored code 2018

The run times in the code significantly decreased. Refactoring the code is important because of the time advantage and efficiencies you can create, especially when working with large datasets. Initially, Steve was looking over only 12 stocks related to Green Energy, but he now wants to use the code to expand it to the remaining stocks in the market. If we did not refactor our code it would take significantly longer to run the analysis on the entire stock market and when we consider this in the realm of work, inefficiencies cost money.

#Code Comparisons: Original vs Refactored Code

The original code contained a nested for loop and within the nested loop it would output the data creating a significantly higher number of iterations.
![This is an image](https://github.com/weise142/stock_analysis/blob/c0d82116ee1692ff8128a306f764beb42b19890a/Initial%20Code.png)

In the refactored code we created one loop, without nested loops, and moved the output of the data to its own separate for loop. This drastically reduced the number of iterations for the macro to run through, thus reducing the run time.
![This is an image](https://github.com/weise142/stock_analysis/blob/c0d82116ee1692ff8128a306f764beb42b19890a/Refactored%20Code%201.png)![This is an image}(https://github.com/weise142/stock_analysis/blob/c0d82116ee1692ff8128a306f764beb42b19890a/Refactored%20Code%202.png)

##Advantages
As noted earlier, the importance of refactoring this code is because it will allow us to run our analysis on a larger dataset in a shorter amount of time. Since Steve is considering running an analysis on all stocks, this means an analysis on over 40,000 stocks on 252 trading days in a year. With that many data points, having efficient code is important. Another impotant factor of refactoring the code is that it becomes more maintainable, meaning if an error occurs or we want to add data points we are better able to troubleshoot and insert that data into our macro. 

##Disadvantages
One of the disadvantages of refactoring our code is that we know our original code worked, meaning we know we can get results when using it. When we attempt to refactor our code we may land in a situation where we have no idea how to proceed or debug an error and unless we are able to eventually figure it our or get help then that was wasted time. We also do not know how long it will take us to refactor code. Since we know we have usable, working code, it means that using additional time to refactor code may not be worth it or if we spend too much time refactoring the code, the time savings we gain in refactoring may may not equate or exceed the time we spent refactoring meaning we made our code less cost efficient. 
