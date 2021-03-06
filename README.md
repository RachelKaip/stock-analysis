# Green Stock Analysis 
## Overview
---
Our friend Steven's parent's are so proud that he just graduated with a degree in Finance and they want to be his first customer. Being passionate about green energry, they found DAQO New Energy Corp (DQ) and ivested all their money into this stock wihtout doing substantial research.  As Steve will be managing thier accounts, he not only wants to look into DQ, but also other alternative energy stocks and potenatially diversify their investments.

Steve created an Excel file with performance data for a handful of green stocks, including DQ, from both 2017 and 2018 and wants us to analyze it.  After reviewing the data set, we set out to find the total daily volume and yearly return of all 12 stocks included in the document using VBA.  We hope that this information will help steve and his parents understand the historical performance of each stock at a glance.  

## Results 
---
### Stock Performance: 2017 vs. 2018 
#### All Stocks
At a glance, stock performed significantly better in 2017 than 2018 with only one stock decreasing in return value.  
![All_Stocks_2017](https://user-images.githubusercontent.com/94569240/147702152-d5b80150-65d8-418f-a8dd-78145373853e.PNG)

![All_Stocks_2018](https://user-images.githubusercontent.com/94569240/147702159-a57e9108-7bb3-40e1-bbc5-f9a1f585f7df.PNG)

Unfortunately, DQ was hit the hardest of the 12 stocks listed, with a -62.6% return value in 2018.  However, ENPH and RUN were the only two stocks to stay in the green in 2018.  In fact, both of these stocks sat comfortably with an 81.92% and 83.96% retrun value respectively.      

Based on our findings, I'd reccomend to Steven that his parents reallocate their investment into ENPH and RUN as these are the only two stocks to remain successful over the two year span in question.  

### Script Performance: Original Script vs. Refactored 
As for the scripts we wrote to find this information, we took two different approaches.  

In the first approach to create the original script, in short, with our nested loop, we instructed Excel to loop through the rows to find every ticker, make it's calculations, and then output that ticker's data before looping again to look for the next ticker.  

![Original_Nested_Loop](https://user-images.githubusercontent.com/94569240/147702393-4a553a97-e444-4806-9fd8-51690d2ddd6a.PNG)



However, in the second approach to refactor the data, one of the main attributes that affects the run times is that we instructed Excel to only loop through the rows once, calculate for each current the Total Daily Volume and Return %, then output all the data at once.  We did this by changing/adding two parts:
1. At the bottom of the inner loop, we told excel that if the ticker in the next row, doesn't match the ticker in the current row, then to "increase the ticker index by 1" AKA, move on to the next ticker in the list before ending the loop's calcualtion.  

=======
However, in the second approach to refactor the data, one of the main attributes that affects the run times is that we instructed Excel to do all the caluculations at once *before outputting anything*.  We did this by changing/adding two pieces of information:
1. At the of the inner loop, we told excel that if the ticker in the next row, doesn't match the ticker in the current row, then to "increase the ticker index by 1" AKA, move on to the next ticker in the list before ending the loop's calcualtion.  
![Refactored_Ticker_Index_Increase](https://user-images.githubusercontent.com/94569240/148440490-9382b8a5-cafa-4bae-a910-743f4134d11b.PNG)

2. The output instructions are in a seperate loop of thier own.  Therefore, Excel couldn't even think about outputting this information before making each calculation instructed in the larger nested loop.  

![Refactored_Output_Loop](https://user-images.githubusercontent.com/94569240/147702551-334cb0a7-80c8-4cdc-b3a2-3a1efd4377d2.PNG)

#### Script Run Times 
In the end, we can see that while each approach accomplishes the same thing, they vary in run times.
---
![VBA_Original_Analysis](https://user-images.githubusercontent.com/94569240/148440777-7273f54f-2224-4858-99fc-954fc926ceb0.PNG)
---
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94569240/148440810-1801f76e-3464-4006-9054-9d61a08afa9d.PNG)
---
As stated in the screenshot above, the original code, where Excel makes and outputs each calculation individually, runs ~2 seconds slower than the refactored script.  I beleive this may be because it's faster for excel to make all the calculations while looping through the rows once than it is to stop the loop to output data before moving onto the next ticker.  

## Summary
---
### Advantages and/or Disadvatages of Refactoring Code
While working through this challenge, I really started to undersatnd one of the beauties of coding- it's flexibility.  While refactoring the All Stocks Analysis code taught me that there really is no "right" way to approach a task, it also taught me that there are advantages and disadvantages to writing and refactoring code.  
#### Possible Advantages
1. Refactoring code gives you the opportunity to streamline the instructions you're giving your machine and still get the same desired outcome- just more efficiently than before.     
2. It also allows you to reformat your code and make it easier to read, digest, and understand.  

#### Possible Disadvatages 
1. Sometimes, streamlining code may require you to add new lines or variables and depending on the instruction given to your machine, this may increase overall run time.  

### How Does this Apply to Refactoring our Original VBA Script?
The refactored version of our script was more logically streamlined (Advantage #1)- meaning that asking excel to make all the calculations first *then* output the results and format the cells would, in theory, be less steps than what the original code instructed.  However,  in order to do this, we had to add an additional If Then statement that *may* have added 2 seconds to the script's run time (Disadvantage #1), making it less efficient in that reguard.   
