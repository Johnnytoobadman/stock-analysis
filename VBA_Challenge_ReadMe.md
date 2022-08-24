# 

## VBA Challenge – All Stocks Analysis

Overview:  Steve is helping his parents in their stock investing as he has some pretty good excel skills.  He has determined that the “green stock” that his parents have invested in is not performing well against a limited set of other green stocks but now he needs to up his game and do the analysis on thousands of green stocks and do so efficiently to better advise his parents on their investing.
	
### Purpose
The purpose of this analysis is to refactor the previous analysis he had done on the limited selection of stocks such that he can efficiently create a data set that easily identifies the stocks with the best returns on an annual basis. Since he is now dealing with a much larger database his script will need to be organized to run as efficiently as possible.

## Results
### Final Product of Steve’s Analysis:

•	For the Year 2017 DQ had the highest return at 199.4%. It also had the lowest volume for the year.  There were a total of 4 stocks with returns in excess of 100%.  TERP had the worst return with a negative 7.2% return.

![](https://github.com/Johnnytoobadman/stock-analysis/blob/main/2017Screenshot.png)![image](https://user-images.githubusercontent.com/110923227/186466234-42ea2da3-329b-4e96-bc01-7c9675b5b048.png)

•	For the year 2018 DQ tanked and had a negative return of 62.6% with 107,873,900 shares traded.  All but two of the selected Green Stocks had negative results with the exception of ENPH (+81.9% and RUN (84%).Both were highly traded. 
•	The run times of Steve’s scripts for the two years 2017 and 2018 was an efficient .1328125 seconds and .125 seconds respectively.

![](https://github.com/Johnnytoobadman/stock-analysis/blob/main/2018screenshot.png) ![image](https://user-images.githubusercontent.com/110923227/186466655-d09c8c4a-5123-4ce4-840a-0b184afdb895.png)

![](https://github.com/Johnnytoobadman/stock-analysis/blob/main/MsgBox%202017.png)![image](https://user-images.githubusercontent.com/110923227/186468041-ac8d75b3-72fb-4c12-9106-011601f90fcf.png)

![](https://github.com/Johnnytoobadman/stock-analysis/blob/main/MsgBox%202018.png)![image](https://user-images.githubusercontent.com/110923227/186467489-4e36c1a0-7545-4bca-a141-3f87cd8ad63f.png)


### Script used to run this efficient subroutine:
•	The initial script from the “All Stocks Analysis (limited set) was reused (refactored) unchanged as below (except subroutine name):



	Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
 
•	Next the tickerindex was created, set to zero and three output arrays were initialized and set to zero: 

  	'1a) Create a ticker Index and set it to zero
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
•	Next a loop was created to loop over all of the rows in the spreadsheet, increase the volume variable for the current ticket inside the loop, check if the current row is the first row with the selected tickerIndex, check if the current row is the last row with the selected tickerIndex and increase the tickerIndex.  This requires using 3 nested if statements within the outer loop.


    
    ''2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Create a Loop that will loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume variable for the current ticker inside the 2b loop
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
•	Next we loop through the arrays to output the results to the worksheet:

	'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    	For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
   	 	Next i
    
 
•	Lastly, we applied formatting, ended the timer, displayed the time in a message box and ended the subroutine in the same code as was refactored:  

	'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

	End Sub

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
