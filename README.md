# Stock Analysis
## Overview of Project
### Purpose
Refactoring the previous code to allow Steve to analyze bigger datasets in a short amount of time.

## Results
### Stock Performance
<img width="230" alt="2017 chart" src="https://user-images.githubusercontent.com/64383146/168412194-2f6dd8c9-a22d-4e11-a79e-ae7bb3a2516f.png">                   <img width="230" alt="2018 chart" src="https://user-images.githubusercontent.com/64383146/168412196-f2504383-58d6-4e7f-8615-c87dfd67a58b.png">

When looking at the 2017 chart, it’s clear because of the colour formatting that this was a good year for a majority of the stocks except one, TERP. The most successful from that year includes DQ, ENPH, FSLR and SEDG. Looking at the 2018 chart and seeing majority red indicates that major drops happened between 2017 and 2018. Ten out of the twelve stocks had their return drop from anywhere between 2 and 260 percent. ENPH dropped in percentage but was still in the green, while the only real successful stock for that year was RUN, increasing their percentage by 78.5 percent. 

### Execution Times
#### 2017
![original_time_2017](https://user-images.githubusercontent.com/64383146/168412214-8e1382df-ba19-46dd-b725-d142c5669302.png) ![VBA_Challenge_2017](https://user-images.githubusercontent.com/64383146/168412216-72576820-34d8-4124-9bee-d9bbceafa379.png)

For 2017, the difference between the original script (left image) and the refactored script (right image) is 0.6166992 seconds.

#### 2018
![original_time_2018](https://user-images.githubusercontent.com/64383146/168412228-68abf134-172e-4b94-8aa1-e9daa5b6043a.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/64383146/168412234-4b43cc64-e788-4f58-8094-b705c4d6f6e4.png)

For 2018, the difference between the original script (left image) and the refactored script (right image) is 0.6130371 seconds.

## Summary
### Refactoring Code
An advantage that comes with refactoring code in general is that when the code needs to be applied to bigger data sets, the refactored code is able to produce the results much faster, making the job more efficient. In addition, when tweaked the code can be used on different datasets, making it more versatile. With that being said, changing the code can be time consuming and sometimes frustrating if the code does not work because of missed details when refactoring it or changing it to fit another data set. 

### Original and Refactored VBA Script 
Looking at the advantages and disadvantages from above, I experienced them while refactoring this code. The advantage of getting the results faster was proven to be true when for both the years the refactored code was able to be about 0.61 seconds faster than the original code. I can imagine that if given a larger dataset, the refactored code would save me time when waiting for the results. I also experienced the disadvantage that comes with refactoring the code, of having to go in and change small details. This allows more room for error, causing frustration when trying to figure out the issue with the code.
