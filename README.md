# Refactor VBA Code and Measure Performance.
## Overview of Project
### Refactoring a VBA code to reduce the execution time of the original code. This code automizes the calculation of the annual volume and the annual return per stock in a selected year. These indicators bring insights about some stocks’ performance and trends. At the same time, this project intends to outline some advantages and disadvantages of code refactoring. 
---
## Results
### Original Code Execution Times 
 (Snipet of Original Code iew.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fraw.githubusercontent.

![2017_Original_Code_Performance](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2017%20Original%20Code%20Performance.png)
![2017_Original_Code_Performance](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2018%20Original%20Code%20Performance.png)

![2017_Output_Original_Code](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2017%20Output%20Original%20Code.png)
![2018_Output_Original_Code](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2018%20Output%20Original%20code.png)

### Refactored Code Execution Times

![2017_Refactored_Code_Performance](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2017%20Refactored%20Code%20Performance.png)
![2018_Refactored_Code_Performance](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2018%20Refactored%20Code%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20Performance.png)

### Code Comparison Example

![Code_Comparison_Example](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/Code%20Comparison%20Example.png)

## Summary
### General Advantages of Refactoring Code
- The code runs faster, which could be critical when dealing with greater amount of data in the future.
- The refactoring changes the structure of the code, making it more organized and logical.
- The code looks more clean and easier to read which is very helpful for other users and future revisits. 
- The code uses less computer memory.
- The code still works if data changes, which prolongs its use in the future.
- **All in all, the code will be more efficient.**
      
### General Disadvantages of Refactoring Code
- The refactoring can be very time-consuming.  That is why it has to be assessed for each project if the time and energy that the refactoring process implies worth it, compared with the expected incremental efficiency of the code execution and subsequent incremental revenues. Also putting in context the importance of the code in the big picture. 
- Refactoring a code does not add a new functionality. Meaning the developer will invest time and effort on something that has been already in use and giving results. He/She might have new projects to focus on at the same time. 
- Another downside is the risk of adding bugs to the original code. That is why is it so important to carefully test the refactored code to make sure it will work, and that takes time too. 

### Advantages and Disadvantages of the Original and Refactored Code. 

**Refactored Code**
- The main advantages of the refactored code in this particular project are: the dramatic reduction of the execution time; as well as having streamlined the code, which also made it more solid and logical. Some of the key edits are:  

        - The use of a tickerIndex.

        - No nested loops but one main loop.

        - The formatting is on the same subroutine


- On a side note, it was observed that the more times the code is run, the faster it tends to run. Please, find these images at the bottom showing the running times registered by the refactored code the first time it was run. It can be appreciated that although execution timings were less than those of the non-refactored code, they were still higher than those shown in the above images, which were reached after running the code several times. 
- As the refactored code runs faster the user of this macro will have the ability to get a faster output, especially beneficial if more tickers are added to the analysis in the future.  
- As for the disadvantages of the refactored code, it still has some limitations, like: hard code (magic numbers) is still used in some loops. For example, those based on the tickers quantity (i=0 to 11). This might need to be changed manually, when new tickers are added, and or others are removed in the future. Even the array with the ticker names has to be manually changed if a ticker changes its name.  
- The refactoring of this code was time consuming. Anyways, if we could have the chance to improve the refactoring, and maybe by using another approach, the macro has the ability to read the tickers from a table, as pivot table would do, and also calculate the total number of tickers in a column, without having to hard code it in the loops.  This would not be new functionality but would add more flexibility and practicality to the code. 

**Original Code**
- The old code has this same limitation above mentioned. Besides, it shows longer execution timings and it is less future-user-friendly.  
- Anyways, the original code still has same functionality than the refactored code. Although, it does not run as fast as the refactored code it has the same fucntionality than the refactored code and it is accurate.
- It still run fast for the size of the data that is being used (*for 12 tickers only*). That said it should not be an issue to use it in this case for the analysis of the stocks’ performance. 

**Outputs** 
- Last but not least, the indicators calculated by the original and the refactored code, together with the conditional formatting used, the user could have an accurate and very straightforward view of the performance of the stocks in 2017 and 2018. 

        * 2017 was a good year for most of the stocks in their category. However, most of them plummetted in 
        
        2018.

        * ENPH and RUN are the only stocks which show positive returns rates in both years. However, ENPH 
        
        recorded a deceleration of the return while RUN registered a dramatic increase from 5.55% to 83.95%.

        * It is recommended to be extremely cautious. A more in-depth analysis of this industry should be 
        
        conducted.

![2017_Refactored_Code_Performance_First_Run](https://github.com/Connectime4ever/stock-analysis/blob/main/Resources/2017%20Refactored%20Code%20Performance%20First%20Run.png)
![2018_Refactored_Code_Performance_First_Run](https://(github.com/Connectime4ever/stock-analysis/blob/main/Resources/2018%20Refactored%20Code%20Performance%20First%20Run.png)