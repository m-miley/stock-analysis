# VBA Code Refactoring

## Overview

This project's objective is to code macros that help analyze the stock reports for all stocks provided and return results to our client, Steve. It is also necessary to refactor the original code so that it will run faster and better accommodate larger sets of data.  Both the original, and the refactored code is included in the Excel document modules.

## Summary

2017 saw a significant increase in Returns for most of our stocks listed with the exception of TERP, and 2018 contrarily saw quite large declines as a whole with the exception of ENPH and RUN.  It seems to me that these stocks are quite volatile and I would not recommend them.  Furthermore, I am not a financial analyst, but Steve is!  So he can take the data we've gleaned from our macros program and present them to his clients for further consideration. 

![2017macros_results](https://user-images.githubusercontent.com/100544761/160896380-de4c1d47-b6a1-40e9-aa89-50cbe445adbb.png) ![2018macros_results](https://user-images.githubusercontent.com/100544761/160896387-284a8be2-26f2-471c-a4bc-7da0a95b225d.png)

## Results

#### Advantages/Disadvantages of refactoring code in general
    - Advantages - improve performance and efficiency.  Also, refactoring can increase the capacity to process varying amounts of data based on the structure of the code.  By including more arrays to hold the calculations as we make them, we can store them and call on them later in the program.  And by inserting new variables such as indexnumber we can better access and update arrays of calculations more efficiently.
    - Disadvantages - spend more time completing a project and maybe not even improving the program all that much.  Depends on the data base size, number of calculations, and whether it's even very important to the final result.

#### Pros and cons of refactoring our VBA script for Steve
  The results of our refactored code were of a major improvement in performance and speed.  The main difference in functionality between the original code and refactored code is the first uses nested for loops to perform the calculations and outputs all the way through, and in the refactored code, by creating 3 new arrays and a tickerIndex, we can store the calculations in arrays by using 3 separate For loops to accomplish the same results. One for loop to set all elements of tickerVolumes array to 0, one main loop to gather and store the info into our arrays, and a third to call the stored information to output.  In doing this, the refactored code ran 4.8x faster on average than the original code.  By using a tickerIndex, we can manage the elements of each array and call them specifically and accordingly.  Also, the upper limit of the tickerIndex is only limited by running out of tickers in the loop (bottom of data table) and therefore easily extendable to the arrays, increasing the number of elements as the number of tickers increase.  Of course, we would have to update this in the array declarations as a variable amount of elements based on something such as input or updating the code as the data grows.

![Original_Code_2017](https://user-images.githubusercontent.com/100544761/160894650-591c4c46-8c75-421e-af3b-7cfb1351473e.png) ![Original_Code_2018](https://user-images.githubusercontent.com/100544761/160894664-1fbe4242-292f-4fa7-bfbe-ec364fdb3af0.png) Original Speed

 ![VBA_Challenge_2017](https://user-images.githubusercontent.com/100544761/160894668-7b79a937-583f-46f4-ae62-eb76e7b33a85.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/100544761/160894682-a1b4d7a5-cc1a-46c3-98a9-a018d3e83fb7.png) Refactored Speed

-------------------------------------------------------------------------------------------------------------

![Code_original](https://user-images.githubusercontent.com/100544761/160899755-a7c9cfdc-5cd9-4606-bcd1-c0c81128ed0d.png) --Original Code segment--

-------------------------------------------------------------------------------------------------------------

![Code_refactored](https://user-images.githubusercontent.com/100544761/160899766-1b7070ba-0b89-4e1f-8541-6bcfc2a2d988.png) --Refactored Code segment--




