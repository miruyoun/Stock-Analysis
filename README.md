# Stock Analysis with VBA

## Overview of Project

### Background and Purpose
For this challenge, I wanted to help Steve be able to analyze stocks faster and more efficiently. Even though the original code I had worked with had a small number of stock variables, it might not work well or be slow with a large number of stock variables. I refactored the original code so that it can loop through the data once and still be able to produce the same output as the original code. The purpose of this task is to show if I was successfully able to make the VBA script run faster with the modified code.

## Results

### Original 2017 Stock Time
![Original_2017](https://user-images.githubusercontent.com/106292020/173379928-37b3c3ed-0de9-460e-8123-6db62b3a7e05.png)
### Original 2018 Stock Time
![Original_2018](https://user-images.githubusercontent.com/106292020/173379983-35f818e6-bef3-4958-a53f-95787943164b.png)
### Refactored 2017 Stock Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/106292020/173380429-764b48ad-d778-47b2-8e8c-7d3064e52caf.png)
### Refactored 2018 Stock Time
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106292020/173380454-ad305f08-fcc2-48de-aef5-0e0c00f42093.png)


My refactored code reduced the time it took for stock analysis to run since the original code sometimes took around 0.75 seconds to run, whereas the original code generally took around 0.1 second. Using "Dim tickerVolumes(12) As Long", "Dim tickerStartingPrices(12) As Single", and "Dim tickerEndingPrices(12) As Single" allowed me to save and recall the output arrays much faster than running through each row and obtaining a value.

Looking at the stock performance using the original and refactored code, both produce the same returns, but it is much easier to figure out which stocks were positive and negative in the refactored code. In 2017, stocks mostly returned positively, while in 2018, the stocks mostly returned negatively.

## Summary

Refactoring is a vital part of any coding process. In general, refactoring allows an improvement of a program's design, structure, or logic while preserving its functionality. Refactoring has the disadvantage of having you look at other people's code, which can be confusing, time-consuming, and potentially introduce more bugs. 

The advantage of my original code was that it was more concise as we did not need to highlight the positive or negative returns. However, the disadvantage of my original code was that it took more time to run. A major disadvantage is that the code might take longer to execute if a significant number of ticker variables are included. In addition to being faster, my refactored code is easier to follow logically. However, it is difficult to follow my code because there are too many variables with the name "ticker".
