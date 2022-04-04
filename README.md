# Stock Market Analysis using Excel VBA

## Project Overview
  This project was undertaken to help our friend, Steve, determine which (if any) of the stocks included in his list would be worthy of investment. Later, the scope widened to include the ability to apply our script to expanded data sets, including more tickers or different years. To complete this analysis, we created a script in VBA for Microsoft Excel that takes a list of stocks tickers and their daily performance, and tabulates them into a single, easy-to-read tabe showing their total performance for a given year. A simple, clearly labeled button prompts Steve to run the analysis, and asks which year he would like to run the script on. Once entered, the table is created and a message box appears letting Steve know how long the program took to run. The performance of the stocks included on Steve's list is included below, as well as a breakdown of the efficiecy of our code.
  
## Results
  ### Stock Results
  We will begin by examing the stock performance extracted from the dataset by running our script. Overall, 2017 was a great year for the stocks in our dataset, with 11 of the 12 tickers experiencing positive growth and half of the stocks seeing returns of at least 50%.  
    ![2017_stock_performance](https://user-images.githubusercontent.com/100643755/161464146-ea9a7ae8-7324-4692-8adb-b2d270818a2a.png)
  
  2018 was a much harsher year, with 10 of the 12 tickers experience negative growth over the course of the year. Of the two positive-growth stocks, ticker ENPH was the strongest performer when taken over the two years, growing 129.5% in 2017 and 81.9% in 2018. 
    
   
![2018_stock_performance](https://user-images.githubusercontent.com/100643755/161464157-48f9c610-7dea-40c5-8bce-bccbac0ab61a.png)

  ### Code Results
   At the outset of this project, Steve only asked us to examine the performance of a single stock ticker, DQ. After looking over those results, he widened the scope to include all the stocks from his original data set. To accomplish this feat, we refactored the code we had created to analyze DQ, making the relevant changes to loop through all of the available tickers and output the results into a new table. While our code accomplished this with zero issues, Steve wondered whether we could apply our code to larger datasets, and if so, whether the performance of the code would suffer. After examining the original script, we aggreed that a sufficiently large data set could cause our code to run over a longer-that-ideal timeframe. Before undertaking any further refactoring, we decided to calculate the time needed for our code to execute, and have included the results below.
   
![runtime_2017](https://user-images.githubusercontent.com/100643755/161464162-3da98c6e-e535-4f00-bd0e-c7628b9920d9.png)
![runtime_2018](https://user-images.githubusercontent.com/100643755/161464166-83be27af-0bfc-4b45-809d-aa7537b656d9.png)

  To make our code more effecient we altered our script so that the program only runs through each row of the dataset a single time. We accomplished this by changing the structure of our **for** loops and printing the results from each row into a dictionary of stock tickers. We ran our programm with these changes and again recorded the time performance of our code.
  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/100643755/161464185-fbce75f7-a7b4-4c9e-8e16-33383a53afe4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/100643755/161464190-c71973a7-a1f3-41db-9df1-49f7815aa446.png)

  Clearly, our refactoring was succesful in creating a more efficient script. The results for either year were provided approximately five times faster than the original code.
 ## Summary
  In conclusion, we (and Steve) have learned the importance of refactoring code. We have learned that the first way is rarely the fastest way, and that it behooves a good programmer to step back and take a wider view of the problem, and code, they are working on. Extrapolating specific problems into more general scenarios allows a programmer to write code that is easier to read, quicker to run, and is applicable to a wider range of data.
  In our specific work for Steve, our original code would have fulfilled his needs completely and taken a totally acceptable amount of time to run. However, we now realize that writing single-use code is inefficient and and a waste of computer resources. Refactoring our code not only allows us to perform analysis on other stock-oriented datasets with ease, but also allows us to ensure that even very large data sets can be parsed in an acceptable time frame. 
    
    
