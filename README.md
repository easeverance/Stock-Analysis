# Stock-Analysis
# Analysis of Stock using VBA
## Overview
  The client Steve is estatic with the workbook that was given to him.  However, in an effort to provide his parents with additional information he would like for us to expand the data to include the entire stock market for the last few years instead of just a few stocks.  As stated before Steve liked the previous workbook but now with the additional information that needs to be included the code needs to be refactored in order to run more efficiently.  
  With refactoring the code we will change it so the code loops through all the data one time in order to collect the same information as before.  The point is to determine if whether refactoring the code successfully made the VBA script run faster.  
  
  ## Results
    There were several steps used to refactor the VBA code that was provided. 
    1. The tickerIndex is set to equal zero before looping over the rows
   
   ![image](https://user-images.githubusercontent.com/90973718/135775752-53dbc9ab-df5a-4f0c-839d-3a5a8c37a69f.png)
   
   2. Arrays were created for tickers, tickerVolumes, tickerStartingPrice and tickerEndingPrice
  
  ![image](https://user-images.githubusercontent.com/90973718/135775762-25b99b9c-b658-421b-b5b9-2d7c6bc73b99.png)
  
    3. The tickerIndex is used to access the stock ticker for tickers, tickersVolumes, tickerStartingPrices and tickerEndingPrices arrays.  This was done by creating a loop and looking for the next row and whether the ticker matched or did not and consequently, increase the tickerIndex. 
    
    
   ![image](https://user-images.githubusercontent.com/90973718/135775834-1e8ed54c-29a5-41bb-991e-edd3b3f92f27.png)
    
    4. The VBA code loops through the arrays to output data "Ticker", "Total Daily Volume" and "Return" columns in the worksheet. 
    
    
   ![image](https://user-images.githubusercontent.com/90973718/135775896-6716f867-3a26-4c34-88d3-14933a006e40.png)
    
    
    I was suprised to see that the refactored code ran slightly slower than the original code. 
    
### Pre-Refactored 2017
    
![image](https://user-images.githubusercontent.com/90973718/135779178-a8b20ee9-bbe1-4766-89bc-da046296e7ac.png)



### Refactored 2017

![image](https://user-images.githubusercontent.com/90973718/135777658-60727f92-fd7e-403b-827b-9c4a90b03d63.png)

### Pre-Refactored 2018

![image](https://user-images.githubusercontent.com/90973718/135779147-17782609-a5e8-45fa-b12b-9b5e897599d3.png)

### Refactored 2018 

![image](https://user-images.githubusercontent.com/90973718/135779224-c3b911dd-d937-45e2-86c1-ab89e5171c66.png)

## Summary
As with most situations in life refactoring of code has it's advantages and disadvantages. 

***Advantages***
Can improve the code
Can assist with better understanding
Can assist with bugs (by making the code simplier in fewer steps)

***Disadvantages***
The code could end up being to big and refactoring it causes more errors
The testing outcome can change

***With This VBA***
The advantage of refactoring the stocks analysis code is that when additional data is added (such as tickers) the program can then run through all the data instead of being subjected to just a few tickers.  




    
