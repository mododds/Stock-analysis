# Module 2: VBA Challenge - Stock Analysis
This repository includes the original module code data as well as the refactored code and final results of the VBA challenge

## Overview
The purpose of this project was to refactor code we used throughout the VBA module and make it more efficient. The macro loops through the rows in a selected sheet (year 2017 and 2018 are included in [this file](https://github.com/mododds/Stock-analysis/blob/main/VBA_Challenge.xlsm)) and output the total volumes of a stock throughout that year, as well as the starting and ending price for the stock that year.

## Results
### Stock Analysis
As a whole, the stock performance over 2017 was fairly good. Of the 12 stocks reviewed, only 1 stock (TERP) yielded a negative return. Both stock volumes and returns were low for 2018. With only 2 stocks (ENPH & RUN) yielding a positive return. Even though DQ had the highest return at the end of 2017, their overall daily volume was low, and they had the largest drop in year-over-year daily volume. ENPH still had a good return and high volume for 2018, RUN is trending upwards in both volume and rate of return. I’ve including an additional graph depicting the year-or-year results [located here](https://github.com/mododds/Stock-analysis/blob/main/Year-over-year_Stock-Analysis.png).

### Macro Analysis
The macro could still be updated to run a bit more smoothly for both [2017](https://github.com/mododds/Stock-analysis/blob/main/VBA_Challenge_2017.png) and [2018](https://github.com/mododds/Stock-analysis/blob/main/VBA_Challenge_2018.png). For example, in order to increase it's reusability, the [hardcoded tickers](https://github.com/mododds/Stock-analysis/blob/main/HardCodedTickers.png) could be removed, and replaced with the tickerIndex. (Keeping in line with the module, if the purpose of this project is to explain to Steve's parents where they should invest their money, I’m not sure Steve's parents are VBA experts, and the code would probably go over their heads....)

## Summary
### Advantages and Disadvantages of Refactoring code
#### Advantages
Refactoring code allows for a lot of advantages including (but not limited to):
- Speed of coding: It allows coders to use existing codes instead of starting from scratch
- Minimize errors: It allows coders to use "tried and true" logic statements that they know already work

#### Disadvantages
Refactoring code can also be problematic for coders for a number of reasons including (but not limited to):
- Compile errors: If names of variables are changed in some parts of the code but not others, this can cause issues
- Debugging:  It can sometimes be difficult to find errors as the code wasn't written from scratch

### How refactoring applies to this project
I actually found this project more difficult than it should have been. I have experience writing and updating VBA codes for "screenscrape" type macros from a previously job role. While the module and virtual sessions went over basic logic statements (if/than, loops, etc.), I felt there was not enough variety of code taught to fully explain the point of adding the tickerIndex variable. While it was great to copy and paste the starting/ending price if/then statements into the new code. I felt like a fresh project may have been more helpful in actually learning and understanding VBA logic and coding.  But it was still fun! 
