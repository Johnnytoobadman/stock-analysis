# 

## VBA Challenge – All Stocks Analysis

Overview:  Steve is helping his parents in their stock investing as he has some pretty good excel skills.  He has determined that the “green stock” that his parents have invested in is not performing well against a limited set of other green stocks but now he needs to up his game and do the analysis on thousands of green stocks and do so efficiently to better advise his parents on their investing.
	
### Purpose
The purpose of this analysis is to refactor the previous analysis he had done on the limited selection of stocks such that he can efficiently create a data set that easily identifies the stocks with the best returns on an annual basis. Since he is now dealing with a much larger database his script will need to be organized to run as efficiently as possible.

## Results
### Final Product of Steve’s Analysis:

•	For the Year 2017 DQ had the highest return at 199.4%. It also had the lowest volume for the year.  There were a total of 4 stocks with returns in excess of 100%.  TERP had the worst return with a negative 7.2% return.


•	For the year 2018 DQ tanked and had a negative return of 62.6% with 107,873,900 shares traded.  All but two of the selected Green Stocks had negative results with the exception of ENPH (+81.9% and RUN (84%).Both were highly traded. 
•	The run times of Steve’s scripts for the two years 2017 and 2018 was an efficient .1328125 seconds and .125 seconds respectively.

 
!(/Users/johnlansberry/Documents/UCSD Files/MODULE 2 VBA/Resources/2018screenshot.png)
 


### Script used to run this efficient subroutine:
•	The initial script from the “All Stocks Analysis (limited set) was reused (refactored) unchanged as below (except subroutine name):

....

	Sub AllStocksAnalysisRefactored()
 
•	Next the tickerindex was created, set to zero and three output arrays were initialized and set to zero: 

  	'1a) Create a ticker Index and set it to zero
•	Next a loop was created to loop over all of the rows in the spreadsheet, increase the volume variable for the current ticket inside the loop, check if the current row is the first row with the selected tickerIndex, check if the current row is the last row with the selected tickerIndex and increase the tickerIndex.  This requires using 3 nested if statements within the outer loop.


•	Next we loop through the arrays to output the results to the worksheet:

	'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
 
•	Lastly, we applied formatting, ended the timer, displayed the time in a message box and ended the subroutine in the same code as was refactored:  

	'Formatting
### Summary


## Advantages and Disadvantages of Refactoring Code in General:

•	The reuse of the original code was very advantageous for saving coding time by reusing the following:

	o	Set start and end time variables
	o	Create the year variable message box
	o	Start the timer
	o	Activate the worksheet
	o	Create the headers
	o	Initialize the array of all tickers
	o	Activate the data worksheet
	o	Get the number of rows to loop over
	o	Apply formatting
	o	End the timer and run the message box with the time displayed.

•	The reuse of the original code was disadvantageous in that one must be very cautious when changing the code to avoid syntax errors that result from changing the code.


## Advantages and Disadvantages of the Original and Refactored Script:

•	The original script was much easier to run as it involved fewer variables and was much easier to avoid syntax errors.

•	The original script was also much less efficient and would be a problem when working with very large data sets.

•	The refactored script was much more efficient and capable of handling large datasets quickly. Likewise it was much more powerful in handling the data with minimal coding.

•	The refactored  script was much more complicated with the additional output arrays and the nested if statements. Additional time was required to test and debug the script due to its complexity. Keeping “ticker” and “tickers” straight was a bit challenging.