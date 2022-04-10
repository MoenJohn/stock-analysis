# Results and Refactor of Stock Analysis
An analysis of green energy company stock data and the results of refactoring the code that analyzed the data

## Overview of Project: 
In order to analyze stock data, I created a VBA script to collect the values from a spreadsheet and output results on a separate sheet. The code was sound and worked, but was not efficient enough to do large datasets. After comparing the results of two different years of stock data, I will detail a refactoring of that code and an analysis on the performance boost it produced.

## Results: 

### Stock performance between 2017 and 2018
Based on the results found by both the original and refactored scripts, **2017 was a much better year** for green energy securities than 2018. 

![VBA_Challenge_2017_results](Resources/VBA_Challenge_2017_Results.png)

Not only were all but one company's yearly returns in 2017 positive, but nearly half had exceeded 100% returns that year. As a whole these green tech companies returned **an average of 67.3%**.

![VBA_Challenge_2018_results](Resources/VBA_Challenge_2018_Results.png)

In 2018, numbers were far worse, with only two companies being in the positive. Nearly half of companies had losses of -20% or worse and an overall **average loss of -8.5%**.

### Refactored Code
The most significant change I made to the code was looping through the stock data only once instead of looping through it for every stock ticker value (a total of 11 times). We can see here in the original code that the ticker to search is updated after the loop completes.

![VBA_Challenge_oldloop](Resources/VBA_Challenge_oldloop.png)

The new code avoids this by changing the ticker value after the last value of a given ticker is reached. This was an easy addition since we already were finding the last available value for a given ticker in order to stop adding the next ticker's values. Note how at the end of the code below the ticker is updated whenever the last value has been read.

![VBA_Challenge_newloop1](Resources/VBA_Challenge_newloop1.png)

**...**

![VBA_Challenge_newloop2](Resources/VBA_Challenge_newloop2.png)

### Execution Times
With these changes in place **significant improvements in speed were produced**. Let's check the differences side by side, with the old value on the left.

![VBA_Challenge_2017_old](Resources/VBA_Challenge_2017_old.png) ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2018_old](Resources/VBA_Challenge_2018_old.png) ![VBA_Challenge_2018](Resources/VBA_Challenge_2018.PNG)

As we can see above the speed is improved by nearly a factor of 10, which neatly corresponds to running the loop 10 less times.


## Summary:

### What are the advantages or disadvantages of refactoring code?
When talking about refactoring code I believe the advantages are clear. One can not only significantly increase the speed in which a program runs, as we saw in this project, but other metrics of code performance as well. Refactoring code can also save on memory usage, which might be a consideration with a project. For example, when saving over an immutable object such as a string over and over again (when building a string in a loop), we fill up memory with each change. Using built in modules and tools can sometimes fix this outright and is an easy refactoring. Another reason to refactor is to improve code readability. After getting some code working it pays to go back and add comments, improve spacing choices, and so on, so you and others can read the code and understand it easier later.

Disadvantages of refactoring code are not as obvious, but a legitimate concern. I would consider one of the biggest disadvantages that it could be a waste of time to refactor code that works fast **enough**, and is memory efficient **enough**. With the power of modern hardware we do not have to be as diligent making code run efficiently, and perhaps some small performance hits are ok for the sake of code readability. You also run the risk of stopping your code from working when you do an ambitious refactor, reminding us of the importance of version control and saving often.

### How do these pros and cons apply to refactoring the original VBA script?
The biggest advantage of the refactor here is the improvements in speed, allowing for much larger datasets to be fed into the program. I would say as far as code readability is concerned, it is a bit of a lateral move. I would say the main advantage of using loops like we had before the refactor is how intuitive they are to read. However it doesn't have much impact here. Finally, if a larger dataset is never fed through the program perhaps we could see the refactoring as a 'time waste', however we could rationalize it all as best practice and stay positive! Ultimately, I think this refactoring leaves us with more useful code and I'm happy to add it to my toolbox in its refactored state over the original.
