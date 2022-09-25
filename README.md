# Stock Analysis

## Overview
In this project, we used Visual Basic for Applications (VBA) to analyze several green energy stocks for the years 2017 and 2018. We investigated  the total daily volume for each stock to measure how actively each stock was traded as well as the yearly return. In our first analysis, we developed code to run this analysis. However, we determined the code may not be efficient or work well for analyzing a larger data set. After successfully calculating the required information, we attempted to refactor the code to see if can maky taking fewer steps, using less memory, and improving logic. 

## Results
After writing code in VBA to analyze stock data for the years 2017 and 2018, we found the following results:
![VBA_Challenge_2017_Results](https://user-images.githubusercontent.com/112137694/192166040-d3075926-94bc-4609-b345-542f3a825209.png)

![VBA_Challenge_2018_Results](https://user-images.githubusercontent.com/112137694/192166042-78063524-9ef3-4ea4-8f32-1cea93652581.png)

As seen in the data above, we can determine that overall 2017 had a higher rate of return than 2018. The only stock in 2017 that did not see a positive return was TERP. In 2018, the only two stocks that saw a positive return were ENPH and RUN. DQ saw a significant decrease in return from the years 2017 to 2018, even though its total daily volume seemed to increase. This contradicts Steve's parent's assumption that the more a stock is traded, the better it will perform. 

### Refactoring
After performing our initial analysis, we decided to refactor our original code in order to enhance the efficiency of our code. The code we currently have written works well with smaller data sets, but if Steve wanted to perform the same analysis on the entire stock market, it would take a siginificant amount of time to perform. 

#### Portions of Old Code
![VBA_Challenge_Old_Code](https://user-images.githubusercontent.com/112137694/192166333-9cfbebc5-c3ac-4084-adaa-6258fca95dcd.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/112137694/192166357-7373df6e-8e78-4e38-86f1-150fd88d415e.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/112137694/192166359-a01dbaf5-a662-492b-a6be-dad12280aa25.png)

In the initial code, we looped over the data set row by row. We used only one array. This process led to an execution time of 1.199219 seconds for the year 2017 and 1.261719 seconds for the year 2018. This isn't a signficant amount of time, however, if the data set were larged, it would take quite some time to perform this opertation. Therefore, we refactored the code to attempt to increase efficiency. 

#### Portions of New Code
![VBA_Challenge_New_Code_1](https://user-images.githubusercontent.com/112137694/192166464-b58c23ec-0caf-4ca7-8ed4-64434f2642d4.png)

In our new code, we created three output arrays and a ticker index which allowed the values for each ticker to be stored based on the index. 

![VBA_Challenge_New_Code_2](https://user-images.githubusercontent.com/112137694/192166469-09d2b486-3c23-4640-b552-fa7d55164723.png)

The refactored code allows us to loop once through the data and output the ticker, total daily volume, and return into their arrays. After the loop through the data is finished, a loop through the arrays will output the data needed by calling on the tickersIndex for each value. 

![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/112137694/192166471-88a52c9b-1d01-4d0a-8eaf-34e130578ae5.png) ![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/112137694/192166475-7ed44d90-2d2b-4d38-9d9b-7509b54455b0.png)

After refactoring our code, we can see that the execution time decreased drastically. The run time for 2017 was 0.2109375 seconds and the run time for 2018 was 0.1796875 seconds. This means that this code will work better for a larger data set if needed in the future. 

## Summary
- What are the advantages and disadvantages of refactoring code?
  - There are many advantages to refactoring code. You can increase the efficiency of the code and enable it to work better with different sizes of data. Additionally, the code can be made to be easier to understand for future individuals working with the code. Additionally, the code can be made more adaptive for long-term goals. Refactoring can also flatten code by removing unneccesarry for loops.
  - Additionally, there are disadvantages to refactoring code as it can be a difficult task. Risk is a disadvantage if the code is not easy to understand. If the code is not clear, it may make it difficult to refactor. If the code is very large, it may be difficult to comb through to make editd and refactor the code 

- How do these pros and cons apply to refactoring the original VBA script?
  - A major pro of refactoring our original VBA script is that the new code is much more efficient and runs faster. Additionally, the data is able to be stored better and called upon later. With the old code, the code had to run through all the data and output the values for each ticker before moving on to the next and rewritting variables. 
  - There weren't really any cons to refactoring our code. The only con would be if someone didn't understand the code and made small errors that were hard to track. 
