# stock-analysis
## Module Two Challenge
### Overview of the Project
Our client, Steve, wanted to easily find the total daily volume and the yearly return for each stock. The total daily volume is simply the total number of shares traded throughout the day, while the yearly return is the percentage difference in price from the beginning of the year to the end of the year. Knowing the yearly return for all the stocks can help Steve make a decision on what stocks to buy. We've allowed Steve to find the total daily volume and yearly return of all the stocks for both the years 2017 and 2018.
### Results
#### 2017 Results
After writing all the code and running the analysis on all the stocks for the years 2017 and 2018, we can conclude a few things. First, it is apparent that the yearly returns for almost all the stocks (except TERP) in 2017 were positive. In the photo below, we can see that the rows colored green all had a positve yearly return, and the one stock in red had a negative yearly return.
![This is an image](https://github.com/RyleeJensen/stock-analysis/blob/main/Resources/All_Stocks_2017.png)

In the next image, we can see an example of the code written in order to format the cells to turn green if the return is positive or turn red if the return is negative.

![This is an image](https://github.com/RyleeJensen/stock-analysis/blob/main/Resources/Example_of_Code.png)

Finally, the following image tell us the execution time it took to run the refactored script for analysis on the year 2017.

![This is an image](https://github.com/RyleeJensen/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### 2018 Results
When we run the analysis on all the stocks for the year 2018, we see a lot more red cells than we did in the year 2017. This tells us that all of the stocks, except ENPH and RUN, did quite poorly. It would probably be a good idea to stay away from investing in any of the stocks that pop up red. We can see an image below of what the analysis looks like. The only two stocks that did well in both the years 2017 and 2018 were ENPH and RUN. These stocks would be worth investing in.

![This is an image](https://github.com/RyleeJensen/stock-analysis/blob/main/Resources/All_Stocks_2018.png)

We can also see in the image below what the execution time was to run the refactored script for analysis on the year 2018.

![This is an image](https://github.com/RyleeJensen/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The execution times on both years are almost exactly the same, however the analysis on the year 2017 was a tiny bit faster. Both of the execution times of the refactored script are significantly faster than the orignal script. Looking at the original script, it takes 0.71875 seconds to run for the year 2017 and 0.7226562 seconds to run for the year 2018. 
