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

    \'2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        'loop for tickers
            For j = 0 To 11
    '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(j) Then
            tickerVolumes(j) = tickerVolumes(j) + Cells(i, 8).Value
            End If
     3b) Check if the current row is the first row with the selected tickerIndex.
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
            End If\

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
## Return Equation Explanation
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
While I was not able to find the solution so that the Tickerindex could be used as the index item for the seperate arrays; I was able to refactor the code from module 2 to arrive at the same chart from the module 2 class work.  And in fact the refactored code was faster than the Module 2 code we had developed from the lesson plan.  The images below display the timed differences between the VBA Challenge Refactored Code & the All Stocks Analysis code built during the Module 2 classwork.  

## VBA Challenge 2017 Timer
![VBA Challenge 2017 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
## All Stocks Analysis 2017 Timer
![AllStocksAnalysis 2017 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/AllStocksAnalysisComparisonTimer2017.png)
## VBA Challenge 2018 Timer
![VBA Challenge 2018 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
## All Stocks Analysis 2018 Timer
![AllStocksAnalysis 2018 Timer](https://github.com/Gkmb2390/stock-analysis/blob/main/Resources/AllStocksAnalysisComparisonTimer2018.png)


## 2017 Report Conclusions
The completed table from the 2017 reporting show that virtually every stock,except TERP, had a positive return on the year; however DQ had the highest return with 199.4%.  Steve's parents should hopefully be rather happy at this point in their investment.

While high returns are most certainly a desired reult when investing in stocks, another significnat consideration is the total daily volume, which DQ was the lowest trader amognst the group trading just under 36 million shares a day. The Total Daily volume is often associated with liquity of the overall stock and how easy it would be to trade, or cash out of. 

## 2018 Report Conclusions
The complete report from 2018 shows rather the opposite of 2017 with almost every company reporting negative returns - DQ reporting the greatest loss at 62.5%. Depending on their view of stock traiding Steve's parents may be regreting their investment at this point, however the stock is still up over 100% of what they may have originally purchased.

While the returns for most companies were less favorable than one might hope, we can see a significant increase in the total daily volume of the DQ stock rising to nearly 108 Million shares being traded daily.  This could be an indicator of significant growth or expansion for the company, which could indicate a strong future.  


### Considerations & Future Inclusions for Reporting
Something to consider in future analysis would be the original stock purchase price & number of shares.  While the nearly 200% return for DQ stock may seem significant - without clarity on the stocks original purchase price it may just be a nice number to show. 

If the purcashe price was $1,000 per share or $10,000, and they have a 100 shares of stock; Steve's parents could be seeing a significant amount of money moving their way.  However if they only invest $5 in 1 share of stock that %200 increase doesn't seem as significant. 

# Summary of VBA Challenge 

## What are the advantages & disadvantages of Refactoring code?

Some of the advantages of refactoring code include:
    1) Time savings for calculations being run  
    2) Fewer coding lines necessary
    3) Less complex structures 

Some of the disadvantages of refactoring code include:
    1) May require more time to understand background of refactored code.
    2) Refactoring code into a complex existing system, may have setbacks for both time and money


## How do these Pros & Cons Apply to Refactoring the original VBA Script?
    We are able to see how some of these advantages apply to the VBA Challenge by the nature of some of the requirements we were expected to achieve.  
### Advantages of refactored code examples
    1) For example we are including screenshots of the comparisons of how much time it took for our Refactored code to run, when compared to our All Stock Analysis code.  In both sets of examples posted above, we cut the processing time nearly in half.  While those time differences are barely noticable on such a small set of items, it could easily be compounded if out data set increased into hundreds of stock tickers or even thousands. Knowing that we could be spending roughly half the time computing the results would be fairly signifcant for both our tool and our ability to report on results. 
    2) Similarly when comparing the number of lines of code being written between the Refactored code and the All Stocks Analysis code the overall lines in the 2 process are only off by about 20 or so lines.  A relatively minisucule difference, most likely won't save too much money or time - when it comes to writing the code. However as in the previous example, as the Code gets more complex the differnce in number of lines between the operations could become very significant  
    3) The intention of refactored code is to make the overall structure less complex - and therefore easier to understand at a glance.  This could be incredibly beneficial for teams that are going through a reorganization interanlly, new hires who are needing to be caught up quickly on a project or projects that need to be handed off to other team members.  If for example I needed to hand off my portion of this project to my next team member, it would be easier for him/her to understand the less complex code. 

### Disadvantages of refactored code examples
    1) While the code is easier to understand at a glance - it may require additional time to understand if the code is being integrated with several other processes in development.  For example if in using refactored code I incorporated a variable or array from a seperate, but integrated subroutine, a new team member may not fully understand how this code is meant to interact with the other subroutines. If in the example of the VBA Challenge, I needed to hand off my portion to a new team member but they had not been given the full background of what we are trying to achieve; it may cause bugs in their updates of the code. 
    2) One of the main considerations you would need to make for including refactored code in your code, should be if it would take longer to find & redesign the code you are considering refactoring or if it may be more time efficient to write the code yourself.  If we hadn't been given a good portion of the code that we used for VBA Challenge, then it may not have been worth our time to do the research to find the coding and integrate the process into a new subroutine - when we had a working subroutine in the All Stocks Analysis
