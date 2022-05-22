# Stock Analysis using VBA + Excel


## Overview of Project

In this project, we edited, or refactored, All stocks Analysis solution code from the green stock database, to loop through all the data at once in order to collect the same information. The purpose of this project is to determined whether refactoring our code successfully made the VBA script run faster.

## Results

The results are divided into two parts. The first part covers the stock analysis between the years 2017 & 2018.
The second part covers the analysis between the original & the refactored code.


###### ***All Stocks Analysis***

In the pictures below dipicted are orignal and refactored 2017 & 2018 all stocks analysis with run time.

<img width="713" alt="Orignal_2017" src="https://user-images.githubusercontent.com/104603128/169671510-5cd0a410-4f0b-41af-9b34-8920468b366a.png">

<img width="706" alt="Orignal_2018" src="https://user-images.githubusercontent.com/104603128/169671494-849303a6-5bb0-4006-9fcc-d5497789ee30.png">

- All stocks produced positive returns in year 2017 except TERP.
- ENPH & RUN are the only stocks with positive returns in year 2018.
- ENPH Total daily volume increased exponentionally from 221,772,100 to 607,473,500 whereas RUN Total daily volumes increased from 267,681,300 to 502,757,100 showing that there was more interest towards these stocks and they produced good returns.
- Rest of the stocks performed poorly giving negative returns with DQ losing the most value.


<img width="571" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104603128/169671053-409ee530-ddb4-4d96-a128-c82b0defa93f.png">
<img width="587" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104603128/169671055-857e5bb2-0e77-4649-afb1-68fd7d9e975f.png">


 
###### ***Orignal & Refactored Code Analysis***

- <img width="338" alt="code-orignal" src="https://user-images.githubusercontent.com/104603128/169672333-11e1c35b-37dc-43ee-be19-9be1dfed5203.png">
-The orignal code ran in 1.125 seconds for the year 2017 where as the refactored code ran in 0.1875 secods for the same year.
-For the year 2018, the orignal code ran for 1.15625 seconds whereas the refactored code ran in 0.1875 seconds.
-In the orignal code we used nested loops which made the code smaller but increased the run time.
-The refactored code used array and conditional loops which have a strong effect on increasing the performance if used correctly.


##Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

##Submission

## References

http://www.cpearson.com/excel/ArraysAndRanges.aspx
