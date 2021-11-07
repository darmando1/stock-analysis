# Stock Analysis
## **Summary**
The objective of this report was to create a script that identifies trends for a select group of stocks in the years 2017 and 2018.
## Code Purpose and Modification
Staff initially created a script that included a for loop within a for loop. Although this provided the correct results, staff realized that the script could be modified due to the ticker symbol being in alphabetical order. At a higher level this meant that staff could run through the entire excel sheet once versus running through the excel sheet once for each ticker. At a more technical level staff breaks down how the code was modified for efficiency...

Previously the staff used

  
       For i = 0 to 11
       ticker = tickers(i)
       
       totalVolume = 0
       
       5) loop through rows in the data 
       Worksheets("2018").Activate
       
       For j = 2 to RowCount
       
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
           
This code would take a single ticker from 0 to 11, go through the entire sheet (code comments 5b and 5c), and calculate the totalVolume, startingPrice and endingPrice. It would then output the data for the current ticker into our "All Stocks Analysis" worksheet and then REPEAT. Thus, the code was repeated a total of 12 times in this example through the entire workbook. Staff modified the code to create one singular for loop for the entire sheet by setting the output as arrays and utilizing a "tickerIndex" variable. 

    '1a) Create a ticker Index
    
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
       
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
         
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
           '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
              
               End If
    
    Next i
    
The above identifies comments 3b and 3c as the script going through the sheet and identifying where the ticker ends (which is in alphabetical order), once the ticker ends the script immediately increases the ticker index to the next ticker and continues. The staff recognizes the results are the same as the previous script but with a faster run time. The staff ran the optimized macro for 2017 data and found the code ran in 0.1171875 seconds. The staff ran the same optimized macro for 2018 and found the code ran in 0.109375 seconds.  Attached below are screenshots.

![VBA Challenge Year 2017](https://github.com/darmando1/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

![VBA Challenge Year 2018](https://github.com/darmando1/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

## Analysis of Stock Performance in 2017 vs 2018
### Year 2017
Staff reviewed performance of 12 tickers within the year 2017 and found that 11 of the 12 tickers had a positive return. The staff tried to identify some sort of pattern and did not find any identifiable pattern due to the limited data (Only Total Daily Volume and Return % available). In 2017 the stock with the highest Return % was "DQ" with a return gain of 199.4%. Meanwhile, the lowest Return % for 2017 was "TERP" with a return loss of -7.2%.
### Year 2018
Staff reviewed performance of the same 12 tickers within the year 2018 and found that only 2 of the 12 tickers had a positive return. Once again the staff tried to identify some sort of pattern but due to the limited data set was unable to identify a pattern. Staff identified that in 2018 ticker "RUN" had a return gain of 84.0% and ticker "DQ" had a return loss of -62.6%.
### Conclusion of Stock Performance Output
There are a plethora of factors that affect stock performance that cannot be measured by the data sets provided. This the primary purpose of this tool would be best used as a calculator to identify one's current portfolio and determine losses and gains. Unfortunately, most online platforms provide this type of functionality thus rendering this calculator obselete.
## Summary of Code Modification
Some advantages of refactoring code allows for a faster run time. This extrapolated over thousands of runs or even larger data sets can significantly decrease processing time. Some disadvantages of refactoring code may be due to the code being refactored for a singular data set. There is a potential that if the data set changes the code may no longer work. Refactored code may make assumptions unique to the singular data set at the time. Due to the large amounts of stocks being traded on a constant basis, refactoring this code is important as the datasets could reach terrabytes worth of data. 
