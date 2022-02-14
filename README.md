# Building a Portfolio of Profitable Green Stocks ##

## Overview of Project:

My good friend Steve, a Finance grad and Excel power user is exploring multiple stocks&#39; performances so that his parents can make a sound investment decision. Steve has gathered up 2 years of financial transactions of 12 green stocks which matches with his parents&#39; initial investment preferences, companies which can produce fossil fuel alternatives. Steve chose to evaluate 2 simple yet effective methods, transaction volume and rate of return of stocks as a starting point.

## Purpose and background:

Since Steve is an Excel power user, he could perhaps develop Excel formulas to perform the key requirements. His concern though is for time efficiency, volume of data to be handled and frequency of exploration for current and future needs, forced him looking into automation with VBA. His ask for results in a click of a button to calculate transactions, yearly rate of return, performance highlighted with positive/negative for any number of stocks in an easily readable format can be tactfully scripted using VBA.

## Results:

My initial response to Steve&#39;s need was using two repeating loops. One loop to search all the rows of the specific field in the worksheet and a next one trying to match the column/field by selected ticker(s) which are stored in an array. Every time a match occurred, required outputs were stored in the system memories which later were printed on the desired worksheet utilizing another loop. The outputs were as desired and produced numbers of transactions and performance by year based on Steve&#39;s input year. Please see output below:

![abc](https://github.com/shamayun/stock-analysis/blob/main/Resources/Execution%20Time%20with%20Single%20Array.png)
Based on current volume of data, Steve was happy with the current output. On the other hand, Steve&#39;s future needs would most likely need exploration of exponential data, more volume. Hence, I decided proactively to refactor the current script which will perform the same task but faster, regardless of data size.

To tackle these evolving needs, I designed the program differently and created predefined space for all the output requirements. Utilized 4 arrays instead of 1 array used in previous method. These predefined memory spaces accommodated all required variables as planned. My refactored scripts looped faster and produced faster outputs. Please see below outputs with new method:
![All Stocks 2017](https://github.com/shamayun/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![All Stocks 2018](https://github.com/shamayun/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary:

While it is common practice to refactor codes, there are advantages and disadvantages. I have listed few advantages, disadvantages and a comparison of both methods below:

![Comparison of Methods](https://github.com/shamayun/stock-analysis/blob/main/Resources/Comparison%20of%20Scripting%20Methods.png)

### Advantages:

• Refactoring improves the Design of Existing Code.

• Design helps to make the code efficient, organized.

• Re-design makes code easily understandable, restructured, and duplicate code is eliminated.

### Disadvantages:

• Refactoring can be expensive and risky when developing bigger applications in the view of management.

• It may introduce bugs.

• Delivery schedule may become very tight.



## The advantages I found with refactoring the stock analysis are:


-The output time was faster and will be able to hand data 115x faster!

-Discovered multiple ways of solving one issue


## Challenges:

-Refactoring was time consuming and almost took 4x longer than initial approach

-Needed to overcome barriers of already conceptualized method and needed to rethink to redesign/restructure existing approach which already generated desired output, except rater inefficient.
