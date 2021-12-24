# stock-analysis
Module 2 VBA work

# Overview of Project
The primary focuse of this challenge focuesed on refactoring our existing code from the Module 2 solution.  

I have called out in the comments, but will restate here, that the Challenge work is located within Module 1 of the VBA Challenge Excel document; while the Module 2 solutions we developed from the course work is located in Module 2 of the VBA Challenge Excel document.

## Purpose 
To refactor our module 2 code and see if we can make the code more efficent and ultimately have it run faster than our module 2 solutions.  

## Analysis & Challenges
This challenge in particular gave me a lot of furstration.  I attempted to refactor the code from the Module 2 solution over serveral different iterations - with no real success.  The majority of the issue stemed from an incomplete understanding of the "tickerindex" and how it was meant to be utilized within the refactored code.  After several more attempts to incorporate the "tickerindex" into the module 2 solution code, I found some marginal success in getting the code to run without error.  I also made the error of misunderstanding that the variables "tickervolumes", "tickerstartingprices" and "tickerendingprices" were meant to be established as arrays - not just dimensions as they were outlined in module 2.

# Analysis of Refactored Code
Since a significant portion of the code was already present in the refactored challenge I will more specifcally cover the areas we were asked to update with our module 2 code.  
### Step 1A
Step 1A is relatively simple as we are asked to create an "tickerindex" variable and that varibale should be equal to zero.  As outlined in the code below:
   '1a) Create a ticker Index
    tickerindex = 0

### Step 1B
Step 1B did give me some issues when attempting to create the arrays.  
I had been misunderstanding the values meant to exist within the () of the each array.  I believed it was meant to be an interchangeable value similiar to (i).  However the error message I recieved made me realize that it required a constant value.  Not being sure what constant value it required, I reviewed the code provided in the challenge; specifically the tickers array - where I understood the value to be 12.  

Step 1B further explained that we would need to establish each of the 3 new arrays: tickervolumes, tickerstartingprices & tickerendingprices with the subsequent data types: Long, Single & Single respectively. 
'1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single


### Step 2A
Step 2a asked for us to establish a loop that would reduce the current tickervolumes to 0 for each of the tickers.  Where this had previously been established in the Modeule work as a single line equation **tickervolumes = 0** this would need to be updated in order to become a loop.  I added the simple for loop outlined below which allowed for each ticker value to be set to 0 before we begin the heavy lifting of accruing data in the next step.

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i


## Nest for Loops for Refactored Analysis
As mentioned above the main issue I found with the refactored code was the inclusion of the "tickerindex" variable.  Since step 3d was not originally detailed in the module 2 solution, it required a fair amount of testing and research to better understand where it should be included.  
The hint for how to increase the volume of the current ticker volumes was not as helpful as I imagined it would be, see code below.  
            *tickervolumes(tickerindex) = tickervolumes(tickerindex) + cells(i,8).value*
Each time I attempted to run that code, I would recieve an overflow error message; so having dedicated several hours to attempting it with that code I improvised a workaround - using code from Module 2 and creating a nested loop for parts 2b - 3d, using the code listed below:

''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    'loop for tickers
        For j = 0 To 11
    '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(j) Then
            tickerVolumes(j) = tickerVolumes(j) + Cells(i, 8).Value
            End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(j) And Cells(i - 1, 1).Value <> tickers(j) Then
            tickerStartingPrices(j) = Cells(i, 6).Value
            End If
        '3c) check if the current row is the last row with the selected tickers
            If Cells(i, 1).Value = tickers(j) And Cells(i + 1, 1).Value <> tickers(j) Then
            tickerEndingPrices(j) = Cells(i, 6).Value
            End If
        '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(j) And Cells(i + 1, 1).Value <> tickers(j) Then
            tickerindex = tickerindex + 1
            End If

## Nested Loop Explanation
The above section of code is where the heavy lifting is happening for the variable ticker calculations.  I will walk through the interactions occuring in each section/subsection below.

### Step 2B
    The 2B code **For i = 2 to Rowcount** as the comment implies, is creating the loop that will cover all the rows in the spreadsheet. Following that line I created a nested loop to cover the ticker variables with a new loop identified as j. **For j = 0 to 11**  This allows for the various tickers to be accounted fo as the "i loop" works through each of the rows. 

### Step 3A
    Moving to 3A we see the calculations for "tickervolumes" array.  As mentioned above I attempted to use the hint within the Challenge instructions, however it did not assist as much as I expected.  I continued to run into overflow errors, even after attempting to update the "tickervolumes" to different data types, such as Single, Double, LongLong, etc.  
    My work around *if then* resulted in the correct values being calculated, so I proceed to adapt the rest of the formulas accordingly.  In the code below it allows for the value of tickers(j)  to be compared against the value found in Cell(i,1). Should those values be be equal it begins adding the values found in the 8th column of the "i" row, adding each consecutive row to the next.
            If Cells(i, 1).Value = tickers(j) Then
            tickerVolumes(j) = tickerVolumes(j) + Cells(i, 8).Value
            End If

### Step 3B
    In 3B similar to 3A though varying slightly, we are comparing the vaules of the Cell (i,1) and that of tickers(j).  Adding a condition, that the cell value before cell(i,1) does not equal the value of tickers(j) that in order for the tickerstartingprices(j) to be equivelant to the cell value of (i,6).  Based on the results of the if then, we can safely assume that if the value of a cell preceeding the cell value, we are currently examining is not found to be euqal to one another we can arrive at the conclusion that we have found the ticker starting price for the ticker (j)
            If Cells(i, 1).Value = tickers(j) And Cells(i - 1, 1).Value <> tickers(j) Then
            tickerStartingPrices(j) = Cells(i, 6).Value
            End if


### Step 3C
    Seciton 3C continues the small changes in code to result in the tickerendingprice(j).  Similar to the previous if then statment, we will also see 2 conditionals that need to be met in order for value of Cell(i,6) to equate to the tickerendingprice(j).  However where previously we were consider both Cell(i,1) and the Cell preceeding it i.e. Cells(i-1,1).value, we adapt the second condition to be the cell following (i,1). 
    When we compare the values found in Cell(1,6) = tickers(j) and the following cell is not found to be equal to the tickers(j) i.e. Cell(i+6,1)<>tickers(j). We can conclude that we have found the final ticker price for the corresponding ticker(j). 

            **If Cells(i, 1).Value = tickers(j) And Cells(i + 1, 1).Value <> tickers(j) Then
            tickerEndingPrices(j) = Cells(i, 6).Value
            End If**

### Step 3D
    The final workhorse for these 2 loops, can be found in 3D which requires us to incerase the value of the tickerindex we orignally established in step 1A.  Using a virtually identical if then statement as in 3C, we can find that if the Cell value of (i,1) = tickers(j) but the following cell(i+1,1) does not equal the same tickers(j) value, we must now increase the value of our tickerindex variable, since we established this variable to be 0 as in step 1A: tickerindex = 0; the solution is simply to add 1 to our value at the time of the loop.  and since the loop will run 12 times as according to loop j (i.e. **For j = 0 to 11**).  Which would allow for each of our tickers to be accounted for.  

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(j) And Cells(i + 1, 1).Value <> tickers(j) Then
            tickerindex = tickerindex + 1
            End If

    We follow the above code by closing out Loop J first with the command Next j, quickly followed by our outer loop i which is closed with the Next i identifier.

## Summary Loop for Arrays

    Finally we arrive at step 4 which will allow for the output of our calculations in the Nested For Loops to be generated.  This code was covered fairly well in our module 2 classwork.  
    
### For Loop for Arrays
    First we establish the loop that will allow for us to work through our ticker array - which as previously established runs from 0 to 11.  Next we ensure that we are outputing our data to the correct sheet, requiring we input the *Worksheets Activate* command.  
        While this technically could exist outside of the for Loop, I took the extra precaution of including it to ensure that all values were output to the same sheet each time the loop ran.  
### Ticker Array Explanation
    Next we need to assign the i values in on both sides of the corresponding array equations.  Since the first cell output is intended to be in Cell(4,1) we must update the value to be Cell(4 + i, 1) which will allow for each unique ticker value to be placed in a new cell according to its array number.  In order to make sure we are determingin unique tickers from the array we must also include the (i) at the end of the tickers(i).
### TickerVolumes Array Explanation
    We follow a similar structure for tickervolumes, adapting the code to Cell(4 +i,2) since we are wanting to move to the next column for output. Again we update the tickervolumes array to include the (i) at the end so we are generating unique values as we continue to move through the array.   
### Return Equation Explanation
    Lastly we move over 1 final column to cell(4+i,3) where we look to output our return value.  The equation being the same from our module 2 classwork, we are able to update the tickerendingprices & tickerstartingprices arrays with the (i) to allow for their unique values to be caluclated in turn as the for loop works through each (i) value. The code ends the loop with the Next i command - completeing the calculations for the assignment. 

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
'Looping through for array values
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i

# Conclusions of the Report

## Refactored Code Functionality & Timing
While I was not able to find the solution so that the Tickerindex could be used as the index item for the seperate arrays; I was able to refactor the code from module 2 to arrive at the same chart from the module 2 class work.  And in fact the refactored code was faster than the Module 2 code we had developed from the lesson plan.  The images below display the timed differences between the AllStockAnalysis code & the Refactored images.  

![VBA Challenge 2017 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![AllStocksAnalysis 2017 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/AllStocksAnalysisComparisonTimer2017.png)

![VBA Challenge 2018 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

![AllStocksAnalysis 2018 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/AllStocksAnalysisComparisonTimer2018.png)










## 2017 Report Conclusions




