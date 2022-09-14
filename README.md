# stock-analysis

## Project Overview
Our good friend Steve has just graduated with his finance degree. His parents want to become his first clients and want to sign him on. Being such great fans of green energy, they want to invest their money into DAQO New Energy Corporation, a company that produces silicon wafer for solar panels. Being the smart lad that Steve is, he thinks it would be in their best interest to diversify their portfolio by analyzing several other green energy stocks, in addition to DAQO stock.
An excel file with stock data has been provided, and Visual Basic for Applications (VBA) has been determined to be the method to analyze the stock data.
VBA is Microsoft’s programming language for all Microsoft Office programs. Although users may not manipulate the main Excel software, they can, however, master the art of programming macros to optimize their time in excel. 

## Results
In the year of 2017, almost all the designated green energy stock options prove to be beneficial by having a net positive return. However, the following year paints a different picture, where almost all the potential stock options have a net negative return, save two (ENPH and RUN).

### Part 1:
-	Create a “tickerIndex” variable and set it equal to 0 before iterating over all the rows. This “tickerIndex” will be used to access the correct index across the four different arrays that will be used: The tickers array and the three output arrays that are created in the second half of this part. The three output arrays that are created are: “tickerVolumes”, “tickerStartingPrices”, and “tickerEndingPrices”. It should be noted that the “tickerVolumes” array should be designated as a long data type and that both “tickerStartingPrices” and “tickerEndingPrices” arrays should be a Single data type.

![Q1_screenshot](https://user-images.githubusercontent.com/111096246/190274367-5eb871d6-1f4e-4d2f-9d6e-6122615ac634.PNG)

Sample code used to fulfilll the requirements for part 1.

### Part 2:
-	A for loop is initialized the “tickerVolumes” array to 0. Afterwards, a for loop that will loop over all the rows in the spreadsheet is created.

![Q2_Screenshot](https://user-images.githubusercontent.com/111096246/190274549-c444693d-168b-4957-a0cf-773b129e5ca4.PNG)

Sample code used to fulfilll the requirements for part 2.

### Part 3:
-	Within the for loop in the second half of part 2, a script that increases the current “tickerVolumes” variable and adds the ticker volume for the current stock ticker, where the “tickerIndex” variable is used as the index. Subsequently, an if-then statement is created to check if the current row is the first row with the selected “tickerIndex”. If it is, then the current starting price is assigned to the “tickerStartingPrices” variable. Afterwards, another if-then statement is introduced to check if the current row is the last row with the selected “tickerIndex”. If it is, then the current closing price is assigned to the “tickerEndingPrices” variable. Lastly, A script that increases the “tickerIndex” is inserted if the next row’s ticker does not match the previous row’s ticker.

![Q3_Screenshot](https://user-images.githubusercontent.com/111096246/190274565-5ffe77c3-e2cc-4870-9f4c-7f2aa47cd9c5.PNG)

Sample code used to fulfilll the requirements for part 3.

### Part 4:
-	Using a for loop to loop through the arrays (“tickers”, “tickerVolumes”, “tickerStartingPrices” and “tickerEndingPrices”) to output the “Ticker”, “Total Daily Volume” and “Return” columns in the spreadsheet.

![Q4_Screenshot](https://user-images.githubusercontent.com/111096246/190274656-d1f43cde-b047-4639-80b0-003ff36a7532.PNG)

Sample code used to fulfilll the requirements for part 4.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111096246/190274694-7b414d3d-c4b3-4949-a270-cf1748d4cb18.PNG)

Screenshot of results of refactored code for stock analysis in the year of 2017.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/111096246/190274700-51c253fc-baf6-44aa-8ba3-95bbb123d0d2.PNG)

Screenshot of results of refactored code for stock analysis in the year of 2018.

## Summary
Code refactoring is a method used to optimize and restructure the current code without adjusting its objective. One of the advantages of refactoring code is that each modification made to the code improves it ever so slightly. Enough of these modifications makes the refactored code much better than the original code. The times in which refactoring should be considered, as per Martin Fowler (Father of Code Smell), should be when
-	Refactoring improves the design of the software
-	Refactoring makes software easier to understand
-	Refactoring helps finding bugs
-	Refactoring helps programming faster

The disadvantages of refactoring are when the software is big and beyond the point of refactoring, or when the developers are not fully aware of what the software’s purpose is.  
