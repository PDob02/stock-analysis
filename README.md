# Stock Analysis Using VBA
**Overview of Project: Explain the purpose of this analysis.** 

The purpose of this analysis was to help our fictional friend calculate stocks for his parents. His parents had been heavily invested in "DQ" since they met at Dairy Queen. We went through many different tickers to find things like ticker signs, returns, starting & ending prices, and volume. During this time we built two seperate codes and measured their performance against one another for the same results. Though we received many "MsgBoxes" we never could determine which version of the code was faster initially. After gaining the necessary assitance in refactoring this project again we were able to determine that the second code was faster. 

The purpose of this analysis was to help our fictional friend calculate stocks for his parents. They had been heavily invested in "DQ" since they met at Dairy Queen. We went through many different tickers to find things like ticker signs, returns, starting & ending prices, and volume. During this time we built two seperate codes and measured their performance against one another for the same results. Though we received many "MsgBoxes" we never could determine which version of the code was faster and will need assitance in refactoring this project again. **Update** After running the data analysis again we can clearly see that the refactored code was more efficient by almost a full second. 

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

**2017**

![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/2017.Analysis.png)

**2018**

![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/2018.Analysis.png)

Summary: In a summary statement, address the following questions.

What are the advantages or disadvantages of refactoring code?

There is a time savings to reusing code. It helps substitute in variables quicker and gives a basic outline of what the code should look like while looking for efficiencies along the way. This was shown by the efficiencies gained below:

**Time Savings** 

**Prior to Refactoring**

![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/2017%20Code%20Timing%20All%20Stocks%20Analysis.png)


![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/2018%20Code%20Timing%20All%20Stock%20Analysis.png)

**After refactoring**

![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/Refactored.Code.2017.timing.png)


![This is an image](https://github.com/PDob02/stock-analysis/blob/main/Challenge_Wall_Street/Resources/Refactored.Code.2018.timing.png)

There was a lot of run time savings with the factored code but it took longer to put into action. One of the disadvatnages of refactoring code is that comments get in the way. The intentions on the first project could have changed over time and it is hard to rememeber why things may have been commented that way. Also, if you did not fully understand the logic the first time committing the code and were instructed to commit it, it is likely you still do not understand it fully yet on your refactored code. Finally, the language we are using is VBA and it does not allow for multiple sheets and Macros on the same PC. This is a hindrance in gathering all the information that one may need in analsysis of the code. 

How do these pros and cons apply to refactoring the original VBA script?

These cons outweighed the pros in my experience since I could not get an output on my first attempt at the refactored code. I think things may have been easier if I just started from scratch with more resources and time at my disposal. Update: After finally revisiting the project I was able to pull my code and find the timings! Many of my mistakes were the loops and tickerIndex variable. This helped give us a fuller picture of the project.

These cons outweighed the pros in my experience since I could not get an output on my first attempt at the refactored code. I think things may have been easier if I just started from scratch with more resources and time at my disposal. ALso, there was only one lengthy example to choose from. Had there been two or more we could have done a greater comparison of the code. 
