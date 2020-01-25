# stock-analysis

## Module 1 in VBA (subroutine name: DQAnalysis)
   Analyzed the DQ stock in the year 2018 based on the data provided. Calculated total daily volume of stocks traded , rate of return for the entire year. This would allow Steve to find the state of DQ stock and to recommend his Dad about his DQ stock investment.\
   ![DQ Analysis for 2018](images/DQStockAnalysis.png)
    
## Module 4 in VBA (subroutine name: DoAnalysisForAllTicker)
   For the same 2018 year added logic to iterate over all the stocks. Calculated total daily volume of stocks traded , rate of return for the entire year. With additional formatting we were able to highlight the stock which has positive returns or negative returns for the year 2018. Since Steve is looking for a better stock for his Dad's portfolio this would help him better.
![All Stock Analysis for 2018](images/AllStockAnalysis_2018.png)

## Challenge
### Module 2 in VBA (subroutine name: RefactoredAllStocksAnalysis)
   Refactored the code to take year as an input and based on the year selected read the corresponding sheet to get the list of tickers and their corresponding startingPrice, endingPrice, totalDailyVolume of stocks traded. Once the data is collected calculated the rate of return for each of the ticker and highlighted positive or negative returns. Also by refactoring we would be able to analyze the data for a given year much faster and make a decision. In our case Steve would be able to look both 2017, 2018 data. Though DQ stock performed really well on 2017 with close to 200% rate of return the subsequent year the rate return was negative with 62% down, while ENPH had profit for both 2017 and 2018 with 129% and 81% respectively. So Steve would be able to recommend ENPH stock for his Dad's portfolio compared to DQ. Look at the attached 2017 and 2018 stock analysis for more details.
   
![All Stock Analysis for 2017](images/AllStockAnalysis_2017.png)
![All Stock Analysis for 2018](images/AllStockAnalysis_2018.png)
