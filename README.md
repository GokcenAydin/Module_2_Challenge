# VBA

## Overview of Project

The goal of this analysis was to expand the size of the stock market dataset by integrating the total daily volume and yearly return values for each stock in 
2017 and 2018 data sheets, as well as to refactor the Module 2 solution code to enable the VBA script run more quickly. 

In the main dataset, the yearly returns were calculated by using the first and last closing prices of the stocks, by which, the percentage of increase or decrease in prices from the 
beginning to the end of the year reflected the yearly return value. The yearly volumes were acquired by summing up the daily volumes of the stocks, and a run button was 
activated in order to run the analysis for any year. 


## Results

Arrays were built to avoid repeats in creating outputs, For loops were used to loop over all the rows in the spreadsheet, If conditionals were used to check the conditionals and give 
values accordingly, and formatting was completed by static and numeric formatting, along with automatically fit column widths, which were already built in the code, the refactoring 
of the code was completed. 

When the analysis run, both 2017 and 2018 stock analysis outputs were confirmed to be same as they were in the module. Visuals - taken before and after the 
refactoring - for the pop-up messages showing elapsed run times for each stock analysis output are included in the Resources file (See the images below).

![VBA_Challenge_2017_first_run](https://user-images.githubusercontent.com/104400293/188531272-0619efe0-5ac2-442e-a4a4-8f5d9e9a1f6b.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/104400293/188531230-6c9586df-b5a5-4719-97e8-ec6e77734f63.png)

![VBA_Challenge_2018_first_run](https://user-images.githubusercontent.com/104400293/188531286-226044f7-87ca-463b-a146-b3d8c772e626.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/104400293/188531302-fc8f0b81-ae77-4684-92ab-c801dbdad5a9.png)

As seen in the visuals, running time for 2017 outputs was "0.0625 seconds" before the refactoring process, and turned out as "0.15625 seconds" after; the same way, while 2018 
outputs ran in "0.078125 seconds" before refactoring, they were "0.1015625 seconds" after. According to these results, the target to attain shorter running time could not be reached in this analysis.

## Summary

In general, refactoring is advantageous for gaining extra time for the analyst. It is a good way of practice to find bugs and debug them. Also a good way to see other methods and gain more perspective.
On the other hand, if it is poorly built, it can take even more time than re-writing it.

In this analysis, refactoring the code was time-efficiently advantageous. Many time-consuming details were already built in the code, and taking a look over it shortly was 
enough to proceed. There was no  errors in the code, and so there was no debugging. As the code was written clearly without errors, there was no disadvantage of refactoring this code during the process. 
