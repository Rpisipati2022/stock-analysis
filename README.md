# Stock Analysis using Refactored Code

## Project Overview:
The initial All Stocks Analysis was a very powerful demonstration of how macros can perform calculations across a data set without one having to physically work on the data set, e.g. pivot it, arrange in alphabetical order, etc.  Leaving the original dataset untouched has the benefit of being able to run the same analysis on different data sets, e.g. monthly updates, etc. 

The purpose of this Refactoring exercise is to get the same results but in a more efficient manner, i.e. with less computing power. It also helps one understand how to use an existing code and adapt it to different t=requirements or to generalize it to run with different data sets.

## Results:
### a)	Single stock analysis:
In the initial analysis we looked at one specific stock (DQ), where we calculated the total volume of stocks trades and the return of this stock (end price – starting price). This showed that, contrary to original perception, DQ was in fact a very poor performer, having lost 60% of its value during 2018!

<img width="538" alt="DQ Analysis" src="https://user-images.githubusercontent.com/99691015/155861889-77cbe4ea-a02a-4f10-892e-1ffd0a71b704.png">

### b) Multi-Stock Analysis:
To help make the decision regarding how one may expand one’s stock picking prowess, we expanded this analysis to look at all 11 stocks in the data set. This was done by creating an array called “Tickers (12)” that contained all 12 ticker symbols. We then iterated across all rows looking for volumes, starting prices and ending prices for each ticker. This was done using IF statements to look for the desired ticker symbol in the dataset: if the condition was met, volumes were added and start and end prices were extracted. We then provided for either 2017 or 2018 data to be analyzed by asking the user to input the desired year. This automated the process somewhat as one didn’t have to make that change in the script. An example for the 2017 and 2018 data is shown below.


<img width="435" alt="2017 All Stocks Analysis Results" src="https://user-images.githubusercontent.com/99691015/155862007-8ade1ea1-dff7-4993-8f05-6ee42fc5b147.png">
<img width="405" alt="2018 All Stocks Analysis Results" src="https://user-images.githubusercontent.com/99691015/155862029-cf49b7c3-a41a-453d-afcb-238e3eaa983d.png">

This shows that 2017 was in fact a pretty decent year for most stocks, but almost all of them took a huge hit in 2018.  

The program used to do this is shown in the Subroutine AllStocksAnalysis()

### c) Refactored Code:
As stated above, the code looped through all 3000+ rows of the worksheet looking for specific stock symbols and conducting analyses based on the outcome. To make the process more efficient, we created individual arrays for each stock’s volume and starting and ending prices and tracked them via a tickerIndex. This speeded up the calculations substantially as we didn’t have to work through the entire data set for each ticker looking to see if it matched the one we wanted. We only calculated the information re. volumes, and starting/ending prices for the specified number of rows that contained all of the required index. This was done by incrementing the tickerIndex as soon as we knew we had moved on to the next ticker – this was when the If statement asking whether we had hot the ending price of a ticker symbol came back “True”. This cut down the number of loops the program had to do to yield the same results in a much faster time.

Examples of analysis for 2018 using the All Stocks Analysis method and the Refactored method are shown below:
<img width="435" alt="2017 All Stocks Analysis Results" src="https://user-images.githubusercontent.com/99691015/155862050-ad18519a-bfa4-4535-9b11-375ef41c8125.png">
<img width="403" alt="2017 Refactored Anslysis Results" src="https://user-images.githubusercontent.com/99691015/155862057-83a318cb-b1a6-4abb-ab9f-dfb3f151eeb7.png">

This shows that both methods returned the same results.
However, the time to execute was substantially different:

Time taken to loop through the data for every ticker:

<img width="272" alt="2018 All Stocks Analysis Run Time" src="https://user-images.githubusercontent.com/99691015/155862079-ec8be725-4f7d-43fa-9d22-870513233e33.png"> <img width="256" alt="2018 Refactored Analysis Run Time" src="https://user-images.githubusercontent.com/99691015/155862084-ccb6eed4-0391-49f2-82dd-a5ba4ead0be3.png">

## Summary
Refactoring Code – General Statements:

#### Advantages:
-	Most of the work has been done and the code needs to be updated to handle slightly different data, maybe the input data format has changed or old data is no longer available and new variables are being introduced into the analyses
-	The efficacy of the analysis has been demonstrated as the company has been using this output for business purposes for some time

#### Disadvantages:
-	The approach that someone in the past took to address this issue is unknown. One has to walk through the programs step by step to understand what was done and why. This may take as much time as it may take to redo the program from scratch.
-	The original data set may no longer be available and new code will have to be developed for the new input variables. The analysis itself may have to change to accommodate this changed inpout.


Refactoring this particular VBA script:
#### Advantages:
-	The general approach to what needed to be done for a specific and subsequently each specific stock symbol was well laid out.
-	How to instruct a computer to look for changes in the data to recognize when on switched from one stock to the next was important to understand – this was shown clearly and the logic was easy to follow.
#### Disadvantages:
-	The approach taken in refactoring where we used arrays for what were earlier plain variables was a little difficult to understand at first. The fundamental difference and the benefits of using this approach were not made clear and one had to guess at what was being asked.
-	The real power of For loops where one used the tickerIndex to move through the data set in a more efficient manner was lost on us initially. 

