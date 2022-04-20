# stock-analysis
Analyzing stocks and their returns

# Overview
## Purpose

The purpose of this project is to analyze the entire stock market over the last few years. Although we have code to do this, we also want to refactor the code in orde to make the analysis more efficient. We will refactor the code by simplifying the code and taking fewer steps than in previous macros.

# Results

As evident by the pictures below, we can see that the code was able to be run in a faster time and became more efficient. Prior to refactoring, it took over 1 second to run the analysis for the stocks in 2017. After the refactoring, the same analysis took less than a second. In fact, the analysis was run 0.15 seconds faster (0.85). It was a similar story for the 2018 analysis. Prior to refactoring, the 2018 analysis took 0.93 seconds. However, after refactoring the code, the code took only 0.85 seconds to run, signifiying a more efficient code.


### Before Refactoring (2017)
<img width="259" alt="VBA_Challenge_2017_Original" src="https://user-images.githubusercontent.com/102189324/164303784-3ee32501-374f-4244-8c99-02e45c6f7166.png">

### After
<img width="168" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/102189324/163904244-df011edc-3f9d-476b-ab83-308cb0361222.png">

### Before Refactoring (2018)
<img width="263" alt="VBA_Challenge_2018_Original" src="https://user-images.githubusercontent.com/102189324/164303799-053eb79c-5c95-424d-a116-ad7f940f9451.png">

### After
<img width="170" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/102189324/163904255-d6ce7e02-3c66-45d0-a9cc-90d0279e1d2a.png">

## What Changed

By adding the tickerIndex variable and setting it to 0, we were able to simplify our for loops.

Before

For i = 0 To 11

        ticker = tickers(i)
        
        totalVolume = 0

After

For i = 0 To 11
    
        tickerVolumes(tickerIndex) = 0
        
As you can see prior to refactoring, we had to initialize a ticker variable AND set a totalVolume variable to 0. Now that we are utilizing new arrays to store data, as well as the tickerIndex variable to help sort these new arrays, we can get rid of these unneccessary steps to help simplify our code and make it more efficient.

Adding one to tickerIndex helps moves us on to the next index to store a new value in after we find the ending price of the stock. Once this happens, we can start the code again.

If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
            

            '3d Increase the tickerIndex.
            
                tickerIndex = tickerIndex + 1
                

# Analysis
## Refactoring Code
### Advantages

The largest advantage of refactoring code is the enhancement of the design of the code. A well designed code becomes easier to follow and thus becomes easier to maintain. With fewer lines of code to go through, it can become simpler to understand what line of code to fix if there is an issue.

### Disadvantages

Code refactoring can also have its disadvantages. When refactoring a code, such as this one, we have a code that is functional. Although our original code may not have been efficient, it got the job done and created an end result that we desired. Howeverr, when refactoring code, and altering code, we could accidentally introduce unwanted errors and bugs. A code that was once working, could become obsolete if these are introduced. This could prove to be very costly if a project has a deadline.

## Application on Original VBA Script

These pros and cons appply to our original VBA Script. Our code has become a lot more simple and has fewer lines of code. This is now a much easier code to navigate. Instead of multiple lines of code potentially having errors, there is now less potential for this occurrence. However, some lines of code become a bit more complex, resulting in potential areas. Once refactoring the code there were multiple occasions when the code would not work they way it was designed to, despite producing quality results prior to refactoring. Uneccessary bugs were created.

The original VBA script worked well, but it was also very long and inefficient.

