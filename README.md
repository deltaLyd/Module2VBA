# VBA of Wall Street - Module 2 Challenge
This file contains a Microsoft Excel Macro-Enabled Workbook, with an associated VBA code, accessed through the "Developer" Excel Add-In on the Ribbon.
The VBA code, also known as a a subroutine, has been designed for my client Steve, so that he can quickly analyze stock (equity) information, specifically Total Daily (Trading) Volume and the Return % for the given year--identified by Ticker symbol--for a variety of stocks, in order to make the best investment choice(s) for his parents.  The Excel file, located here: [VBA_Challenge.xlsm ](https://github.com/deltaLyd/Module2VBA/blob/main/VBA_Challenge.xlsm) contains stock data for 12 companies for years 2017 and 2018.  The original code for this analysis has been refactored in a way that allows it to run more quickly and is now able to be applied to a larger potential dataset.
##Analysis of Returns: 2017 vs 2018
As is clearly highlighted using the programmed Conidtional Formatting, 2017 was a year of strong performance for almost all the stocks analyzed.  In fact, only one company, TerraForm Power Inc (ticker symbol "TERP" in the Excel file, had a negative return for the year.













*I would have altered the Conditional formatting to be the red, yellow, green color-scale, rather than the binary red & green, as this makes it hard to quickly differentiate between moderately successful stocks and highly successful ones: in 2017's analyssis both "RUN" and "DQ" had positive returns for the year, and were highlighted green by the Conditional Formatting. However, RUN returned 5.5%, whereas DQ returned 199.4%, clearly making the latter the better investment. The code should be altered to reflect the difference in performance more clearly, even if both are positve.
