# VBA of Wall Street - Module 2 Challenge
This file contains a Microsoft Excel Macro-Enabled Workbook, with an associated VBA code, accessed through the "Developer" Excel Add-In on the Ribbon.
The VBA code, also known as a a subroutine, has been designed for my client Steve, so that he can quickly analyze stock (equity) information, specifically Total Daily (Trading) Volume and the Return % for the given year--identified by Ticker symbol--for a variety of stocks, in order to make the best investment choice(s) for his parents.  
The final Excel file, located here: [VBA_Challenge.xlsm ](https://github.com/deltaLyd/Module2VBA/blob/main/VBA_Challenge.xlsm) contains stock data for 12 companies for years 2017 and 2018.  The original code for this analysis has been refactored, allowing it to run more quickly and to be capable of execution on a larger potential dataset.  The working Excel file, located here: [green_stocks data file - working file.xlsm](https://github.com/deltaLyd/Module2VBA/blob/main/green_stocks%20data%20file%20-%20working%20file.xlsm) produces the same output, but uses a longer, less efficient code that executes more slowly than the refactored version.  Please refer to the VBA code in both files by using the "Developer" Excel Add-In on the Ribbon.  You will note that the original, unfactored code contains multiple subroutines and is not as concise as the refactored version, running approximately 4 pages to the refactored file's 2 pages.
## Analysis of Returns: 2017 vs 2018
As is clearly highlighted using the programmed Conditional Formatting, 2017 was a year of strong performance for almost all the stocks analyzed.  In fact, only one company, TerraForm Power Inc (ticker symbol "TERP" in the Excel file, had a negative return for the year. 
### Exhibit 1: All Stocks Analysis (Refactored) Output for 2017
![VBA_Challenge_2017.PNG](https://github.com/deltaLyd/Module2VBA/blob/main/Resources/VBA_Challenge_2017.PNG)

On the other hand, 2018 was a disappointing year for stock performance, at least for the 12 tickers I analyzed for Steve. While 2 of 12 stocks provided very strong (80%+ returns), the other 10 stocks all had negative returns for the year.

### Exhibit 2: All Stocks Analysis (Refactored) Output for 2018
![VBA_Challenge_2018.PNG](https://github.com/deltaLyd/Module2VBA/blob/main/Resources/VBA_Challenge_2018.PNG)

## Comparison of Original Code to Refactored Code
While the code for both subroutines produced the same outcome, the refactored code performs the same tasks at a greater rate of speed, and is therefore a better end-product to deliver to Steve, so that he does not have to sit around waiting for the original code to spin through a very large dataset he is trying to analyze. 
Compare the following Exhibit (3) to Exhibit 1. Note that the output is the same, but the time to completion is much faster in Exhibit 1 at ~0.09 seconds, which ran using refactored code, versus the ~0.72 seconds in Exhibit 3, which ran using the original code.
### Exhibit 3: All Stocks Analysis (Original) Output for 2017
![VBA_Challenge_2017 - Initial.PNG](https://github.com/deltaLyd/Module2VBA/blob/main/Resources/VBA_Challenge_2017%20-%20Initial.PNG)
### Exhibit 4: All Stocks Analysis (Original) Output for 2018
Similarily, compare the following Exhibit (4) to Exhibit 2. Again, note that the output is the same, but the time to completion is much faster in Exhibit 2 at ~0.09 seconds, which ran using refactored code, versus the ~0.7 seconds in Exhibit 4, which ran using the original code.

![VBA_Challenge_2018 - Initial.PNG](https://github.com/deltaLyd/Module2VBA/blob/main/Resources/VBA_Challenge_2018%20-%20Initial.PNG)

### *As stated in the introduction, please refer to the VBA code in both files by using the "Developer" Excel Add-In on the Ribbon.  While both codes produce the same output, the refactored version is much neater and requires fewer lines of code.

## Advantages & Disadvantages of Refactoring Code in General
There are many reasons why a programmer should at least consider refactoring their code, though there are a few notable downsides.
Advantages:
-Eliminates redundancies that may have been initially programmed.
-Opportunity to clean up code and comments to make the program easier to read and audit.
-Enables programmer to simplify long or complex sections of code into shorter, more manageable pieces of code.
-Improves performance of the code (as evidenced by the improved execution time in this Challenge).

Disadvantages:
-Takes more time and effort, which means additional expense and opportunity cost to a business.
-Creates the chance to accidentally write an error into the syntax, which really defeats the point of refactoring (making the code better).

## Advantages & Disadvantages of Refactoring this VBA Script
The advantages of refactoring this VBA script are largely the same as the "in general" section above: run-time is improved, the code is simplified, and it is neater to read.

The main disadvantage to the refactored version is that it does not have all the comments I originally built into my un-factored code, which helped me understand the navigation of the program and what individual sections of code were doing.  As this is not an overly long or complex build, comprehension does not suffer too much, but if the code had been longer or more complex, losing those initial build notes would make the code harder to understand or repurpose in the future to use in another program.

### Additional Refactoring Suggestion
*I would have altered the Conditional formatting to be the red, yellow, green color-scale, rather than the binary red & green, as this makes it hard for me as the programmer and tester, and Steve as the end-user to quickly differentiate between moderately successful stocks and highly successful ones: in 2017's analyssis both "RUN" and "DQ" had positive returns for the year, and were highlighted green by the Conditional Formatting. However, RUN returned 5.5%, whereas DQ returned 199.4%, clearly making the latter the better investment. The code should be altered to reflect the difference in performance more clearly, even if both are positve, negative, etc.
