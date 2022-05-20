# Stock-Analysis


## Overview of Project

1. Create VBA macro. 
2. Macros will generate pop-ups, and formats cell values. 
3. Loops and conditions to direct logical flow 
4. Create pivot table and line chart for theatre outcomes. 
5. Run macros to ensure syntax recollection and pattern recognition. 


## Purpose

  The purpose of the analysis is to be able to assess the performance of each ticker for years 2017 and 2018. With the use of VBA run buttons, the use of year as an input will generate an analysis for each stock ticker, using daily volume and stock return % as its outputs. On the background, this was done by generating VBA macros within a worksheet. This includes the use of loops and conditionals to guide the flow of information and using the macro to generate cell values and format cells based on the year.


## Analysis 


### Analysis of Stock Performance between 2017 & 2018

  By comparing 2017 vs. 2018 stocks performance, it is evident that the ENPH and RUN tickers would have been considered good investments. This is because of their positive and significantly high returns in 2018. In addition, both tickers had increases greater than $200,000,000 over the 2017 total daily volumes.
  


## Analysis of Execution Times of the Original Script and the Refactored Script.

  Based on the alignment of the code, the refactored script contributed to a run time decrease in comparison to the original script.
                                  
  ![This is an image](https://github.com/Stookhy/Stock-Analysis/blob/main/VBA_Challenge_2017_Elapsed_Run_Time.png?raw=true')
  ![This is an image](https://github.com/Stookhy/Stock-Analysis/blob/main/VBA_Challenge_2018_Elapsed_Run_Time.png?raw=true')  
    
                                  
  ![This is an image](https://github.com/Stookhy/Stock-Analysis/blob/main/VBA_Challenge_2017_Elapsed_Run_Time_Refactored.png?raw=true')
  ![This is an image](https://github.com/Stookhy/Stock-Analysis/blob/main/VBA_Challenge_2018_Elapsed_Run_Time_Refactored.png?raw=true')
 

## Summary

  1.	What are the advantages or disadvantages of refactoring code? 

  The advantages related to refactoring code include:
    •	Ease in reviewing the work and finding the root causes to debug
    •	Ability to make the code more readable and easier for a user to understand
    
    
  The disadvantages related to refactoring code include:
  •	Requirement to run multiple tests to ensure that the programming code overall       is working correctly
  •	Possibility of creating new error messages or bugs in the overall coding document
  •	Can take a lot of time to trace the errors and debug information
  
  
  2.	How do these pros and cons apply to refactoring the original VBA script?
  
  The pros related to refactoring VBA script include:
  •	Reducing number of loops, which decreases the speed required to run the code  and optimizes its performance. This is evident in the 2017 and 2018 information,  where the speed time when refactoring code decreases as follows
  •	2017 speed time 
      From 3.71 seconds to 2.00 seconds
      
  •	2018
      From 3.60 seconds to 1.92 seconds
      
  The cons relating to refactoring VBA script include:
  •	Running into error messages related to syntax alignment and other formatting requirements
•	As a result, this process can be a time-consuming exercise
