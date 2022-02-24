# Stock Analysis

## Table of Contents
- [Overview of Project](#OverviewProject)
  * [Background](#Background)
  * [Purpose](#purpose)
- [Results](#results)
  * [Stock Performance](#stockper)
  * [Code Performance](#codeper)
- [Summary](#summary)
- [References](#references)

## <a name="OverviewProject"></a>Overview of Project

### <a name="background"></a>Background

Steve, a financier, is helping his clients who want to invest in DAQO New Energy Corp (DQ). He is wanting to analyze a handful of other green energy stocks to see if there's a better stock to invest in. 

Prior to this analysis a macro was prepared for Steve to run this analysis, yet as part of a faster code will be provided. 

### <a name="purpose"></a>Purpose

The purpose of this analysis is to determine which would be the best stock to purchase.  Additionally, the refactored code will be compared to the original code to better understand its benefits and pitfalls. 

## <a name="results"></a>Results

### <a name="stockper"></a>Stock Performance

From the analyzed data we found that 2017 was a much better year for trading these set of stocks than 2018. In 2017, eleven out of the twelve stocks increased in price. In 2018 only two of the stocks increased in price, and all of the other stocks decreased in price. 

The two stocks that increased in price both years were ENPH and RUN. ENPH increased in price by 129% in 2017 and 81% in 2018. Where as RUN had a 5% increase in 2017 and an 84% increase in 2018. 

<p align="center"> <img src="Resources/VBA_Challenge_2017.png" width ="30%" alt="VBA_Challenge_2017"> </p>
<p align="center"> Figure 1: Refactored Code Outcome for 2017</p> 

<p align="center"> <img src="Resources/VBA_Challenge_2018.png" width ="30%" alt="VBA_Challenge_2018"> </p>
<p align="center"> Figure 2: Refactored Code Outcome for 2018</p> 

### <a name="codeper"></a>Code Performance

In order to better serve the clients, the speed the code uses should be taken into consideration. In the original code, embeded loops were used. The outside loop went though all 12 different stocks. The inside loop evaluated all of the rows in the code. There are about 3,000 rows in our data sets for both the 2017 year and the 2018 year. In this nested loop, the code evaluates all of the rows for every different stock, so that means there are aproximately 36,000 iterations. 

Our new code uses a conditional statement to evaluate if the data is moving into the next set of stocks rather than a for loop. Therefore passing though each row only once. The code has only approximately 3,000 iterations. 

The resutls themselves are not different but the time it took the computer to evaluate the results was reduced from __ to __ . 

<p align="center"> <img src="Resources/Module_2017.png" width ="30%" alt="Module_2017"> </p>
<p align="center"> Figure 3: Embeded Loop Code Outcome for 2017</p> 

<p align="center"> <img src="Resources/Module_2018.png" width ="30%" alt="Module_2018"> </p>
<p align="center"> Figure 4: Embeded Loop Outcome for 2018</p> 

I compressed the code even further by creating nested if statements in order to see if this would cut down the time, the code can be seen below. 
...

        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
            ElseIf Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
                tickerIndex = tickerIndex + 1
                
                ElseIf Cells(i, 1).Value = tickers(tickerIndex) Then
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
                    
        End If          
...

Yet the time it took the computer to work through the nested if statement was slighlty more than what it took to work through individual if statements. So the code was reversed to individual if statements. 

...

        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        End If       
      
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
               
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If

...

The way all of these codes are written assumes that the data has been order such that all stocks are grouped together. A potential improvement for future macros would be to include a task that ensures that all of the data is organized by stocks. 

## <a name="summary"></a>Summary

As part of the data analysis we can conclude that investing in ENPH could bring better returns in the long run than investing in DQ. 

The refactored code reduced the iterations from 36,000 to 3,000 making the code faster by a factor of 10 without changing the results. 




## <a name="references"></a> References

[1] [Stock Analysis Excel File](https://github.com/tamiespinosa/stock-analysis/blob/a36556cee6e784b0aa7973acf9afcac611f73115/VBA_Challenge.xlsm)

[2] https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax
