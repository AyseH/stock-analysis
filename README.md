#Green Stocks Analysis with VBA

##Overview of the Project

Steve wants to help his parents in their investment plans in green energy. They are passionate about green energy but they have not done sufficient research about green stocks. Even though they are particularly interested on DAQO Energy Corp (DQ), Steve wants to help his parents by also analyzing other green stocks in the market to enable them diversify their funds.

###Purpose of the Analysis

-The purpose of this project is to provide Steve with a complete report of the green stocks in the market by determining their total daily volume and yearly return for each of them. 
-The daily volume is used as an indicator of the value of the stock by measuring how actively it is traded in the market. 
-The yearly return is used as an indicator to determine how much an investment grew in a given year. 
-In addition to the analysis of green stocks, the project aims to provide Steve with a refactored, more efficient code that will enable him to perform this analysis in a larger data set in a shorter amount time.

##Results

-There are 12 green stocks in our data.
-The results of the analysis of the green stocks by using VBA show that DQ is not a wise choice. Even though, the yearly return of DQ was impressive in 2017, it had a negative yearly return in 2018.
-In 2017, although the yearly return of DQ was the highest of all the stocks, the total daily volume of DQ was the lowest. However, in 2018, the total daily volume of DQ more than tripled while the yearly return dropped to -62.6% in 2018 from 199.4% in 2017.
-When we investigate the performance of other stocks we see that the only stocks that had positive yearly returns in both 2017 and 2018 were RUN and ENPH. 
-The yearly return of ENPH decreased to 81.9% in 2018, from 129.5% in 2017; while it increased to 84% from 5.5% for RUN.
-The total daily volume of both ENPH and RUN increased in 2018.
-TERP is the only stock that had negative returns in both years.
-The data shows that 2017 was a relatively good year for green stocks. All the stocks except TERP had positive annual returns with large total daily volumes.
-The opposite could be stated for 2018, in which most of the stocks had negative returns even though they had high total daily volumes.
-The pattern between total daily volume and yearly return in 2017 and 2018 is not consistent. For some stocks, total daily volume and yearly return decrease (AY, CSIQ, FSLR, JKS, SPWR) or increase together (ENPH, RUN, TERP) while for others even though the total daily volume increases the yearly return decreases (DQ, HASI, SEDG, VSLR).
-The overall results of the green stocks analysis indicate that ENPH and RUN are the best options for Steve’s parents.
-The report also shows that the code was refactored to perform the analysis in a shorter time frame. 
-In the original script, to determine the total daily volume and the yearly return the steps had to be repeated for each ticker. Even though, an array for the tickers were formed in the original script the program still had to repeat the steps 12 times to gather all the data required.
-In the refactored script, we enhanced the coding so that we can loop through all the data once and collect all of the desired information in this step.
-The refactored script improved the execution time and allowed us to perform the same task within a shorter time frame. 
-The images below delineate the improvement in the execution time.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81400525/117375342-66eaf880-ae94-11eb-84c7-ac628571453b.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/81400525/117375346-694d5280-ae94-11eb-9012-c7e0f4426bd9.png)

-The original script ran the 2017 data in 1.0391 seconds, while the refactored script ran it in 0.1758 seconds.
-Similarly, the original script ran the 2018 data in 1.0586 seconds, while the refactored script ran it in 0.1680 seconds.
-Even though, we are only talking about seconds, the impact could be much bigger in a much greater data.

##Summary

###The Advantages and Disadvantages of Refactoring a Code

####Advantages

-Refactoring a code makes it cleaner and more simple. Therefore, it increases the efficiency and improves system performance.
-It reduces technical debt by saving on costs, efforts and time needed to detect bugs and problems.
-Refactoring can also be useful in improving and updating existing products to meet customer demands and needs(https://langate.com/code-refactoring).

####Disadvantages

-It can be too expensive both in terms of money and time.
-If required changes are of a too large scale, then instead of refactoring building a new software should be considered (https://langate.com/code-refactoring).

###The Advantages and Disadvantages of the Refactored Green Stocks Analysis Code

-For our special case, refactoring the code for the green stocks analysis resulted in a much shorter execution time.
-This could be an important advantage if the code is used in large datasets. 
-On the other hand, for this particular dataset which is not very big, the time and effort spent refactoring the code outweighs the benefits.

