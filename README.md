
# Stock Analysis Using VBA | Overview of Refactoring Code

I was originally asked to do an analysis of stock performance for Steve, a recent college finance graduate who wants to help his parents make sound stock investments. Using VBA to write macros, I performed an audit of stock data of a set of 12 stocks over 2017 and 2018. 

After the initial analysis was completed, Steve asked me to expand the analysis to include the entire stock market. I refactored the code to see if I could make it more efficient, so analyzing a larger data set would run quicker.

## Results
### Stock Analysis 
Steve's parents invested solely in DQ stock, and he was correct in being concerned about the diversity of their portfolio. DQ stock, while it had the highest return in 2017 of the analyzed stocks, it took the biggest loss from one year to the next. Almost all of the green energy stocks reviewed in 2018 were down from the previous year, but Steve's parents would likely benefit from spreading their investment over stocks that were less volatile, such as ENPH and RUN. The RUN stock was the only one to increase its return in 2018 over 2017, and ENPH while its returns were down, still showed positive returns in 2018. 

### Code Efficiency 
Another aspect of this analysis was refactoring code and understanding the benefits and determents to doing so. What I discovered after refactoring the code, is that while the time of the scripts varied on each run, the refactored code ran faster than the original code for both 2017 and 2018 in every case. The screenshots below, show those results. 


**2017 Original Runtime** 

![2017 Original Runtime](https://github.com/jmmadson/stock-analysis/blob/main/Resources/Original_2017RunTime.png)

**2017 Refactored Runtime**

![2017 Refactored Runtime](https://github.com/jmmadson/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

**2018 Original Runtime**

![2018 Original Runtime](https://github.com/jmmadson/stock-analysis/blob/main/Resources/OriginalScript_2018Runtime.png)

**2018 Refactored Runtime** 

![2018 Refactored Runtime](https://github.com/jmmadson/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The difference in time it took to run the macro, was due to the difference in how we looped through and stored the data. In our refactored code, we looped through the data set one time and stored our data in arrays. This was more efficient than the [Original Code](https://github.com/jmmadson/stock-analysis/blob/main/Resources/Original_VBA_Analysis_Script.txt). 

While the time to run the macro was only slightly faster in the refactored code, that time will make a significant difference when analyzing a much larger data set, such as all stock data. 

## Summary 

Refactoring code has the following advantages: 
- Can create efficiencies to allow macros to run faster
- When refactoring, putting in comments to organize the code will make it easier for future adjustments

It also has the following disadvantages: 
- It can take more time to develop, than it could save on the actual running of the code 
- You open the door for new bugs and breaking code that was already functioning

In this instance the main advantage of refactoring the original stock analysis code, was the speed we gained by writing more efficient macros. As an additional advantage refactoring helped me better understand the code, learn additional skills in the way of storing data in arrays, and get a better overall feel for the fundamentals of programming. 

The disadvantage of refactoring the original stock analysis code was that it took a long time to get it working, as I created many bugs along the way to work out. 
