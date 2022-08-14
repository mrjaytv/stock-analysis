# Stock Analysis: Green Energy

## Background
A friend, Steve, is helping his parents to analysis handful green energy investments so this will help them make an informed decisons with their investment stocks and ease any concerns Steve has in diversifying their funds. Currently, Steve's parents invested in Daqo New Energy Corp, a company that makes silicon wafers for solar panels. The ticker symbol for Daqo is DQ. Steve has obtain and shared the stock data analysis for 2017 and 2018. In order to help Steve analyze the data efficiently, we are using VBA to automate tasks.

## Overview and Purpose of Project 
The purpose of this project is to assist Steve in analyzing the stock data yearly return and total daily volume while refactoring the original code received from previous programmer so the analysis is better structured and able to perform at optimized efficiency. This would allow Steve to reuse the refactored code for other years of any stock and reduce the chance of accidents and errors.

## An Observation of the Data
One analysis in observing the difference of total daily volume between 2017 and 2018 how many of the tickers yearly return percentage declined in 2018. When initially comparing only the total daily volume numbers in certain tickers you can assume the stock is doing fairly well, such as ticker: **DQ**. However, when the yearly return is added, it's clearer to see that the performance for **DQ** isn't performing as well as we hoped. Another observation is that ticker stock **

|       2017         |       2018         |     
|:---: |:---: |
|![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2017_dataresults.png)|![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2018_dataresults.png)|


## Results from the Refactoring the Code

Refactoring the initial code helped the data anaylsis run faster. In the long run, Steve could add more stock data over time and leverage the refactored code to instead of the original code. Also, as time goes on, it would be beneficial for Steve to continue to have a developer review the code to see if it can be refactored further. The run times is noticeably different. When the initial code is ran, it takes about .6 seconds where as the refactor code takes around .07 seconds for both 2017 and 2018. 


### Model 1: Run times for the original code took around .6 seconds.

|       2017         |       2018         |     
|:---: |:---: |
|![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2017_msgbx.png) |![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2018_msgbx.png) |

### Model 2: Run times for the refactored code took around .07 seconds. 

|       2017         |       2018         |     
|:---: |:---: |
|![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2017_msgbx_refactor.png) |![This is an image](https://github.com/mrjaytv/stock-analysis/blob/5d00896dfb60c9fbb00ff8c73251d4637ee854a2/Resources/2018_msgbx_refactor.png)|



## Summary 

### What are the advantages or disadvantages of refactoring code?

One of the key compenents when refactoring is to understand what the original code is doing and what the end result should be. If the original code does not have any comments or the programmer has a different logic process, refactoring the code could be challenging and could led to greater issues such as introduce new bugs or costly effort. 

On the other hand, refactoring often helps clean up the code by deleting duplicates, divides a big chunk of codes into several methods, and making it easier to understand. This could help programmers resolves bugs and enhance the execution of code. 

### How do these pros and cons apply to refactoring the original VBA script?

When refactoring the original script for this project, the loops were modified and the worksheet was activate once instead of twice which decreased the run time and optimized the performance of the script.  


Source: https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/ 
