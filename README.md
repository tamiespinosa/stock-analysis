# Stock Analysis

## Table of Contents
- [Overview of Project](#OverviewProject)
  * [Purpose](#purpose)
  * [Background](#Background)
- [Results](#results)
- [Summary](#summary)
- [References](#references)

## <a name="OverviewProject"></a>Overview of Project
### <a name="purpose"></a>Purpose
### <a name="background"></a>Background

## <a name="results"></a>Results
<p align="center"> <img src="Resources/VBA_Challenge_2017.png" width ="30%" alt="VBA_Challenge_2017"> </p>
<p align="center"> Figure 1: Refracted Code Outcome for 2017</p> 

<p align="center"> <img src="Resources/VBA_Challenge_2018.png" width ="30%" alt="VBA_Challenge_2018"> </p>
<p align="center"> Figure 2: Refracted Code Outcome for 2018</p> 

<p align="center"> <img src="Resources/Module_2017.png" width ="30%" alt="Module_2017"> </p>
<p align="center"> Figure 3: Embeded Loop Code Outcome for 2017</p> 

<p align="center"> <img src="Resources/Module_2018.png" width ="30%" alt="Module_2018"> </p>
<p align="center"> Figure 4: Embeded Loop Outcome for 2018</p> 

...
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
            ElseIf Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
                
                ElseIf Cells(i, 1).Value = tickers(tickerIndex) Then
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
                    
        End If
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
...

## <a name="summary"></a>Summary

## <a name="references"></a> References

[1] [Stock Analysis Excel File](https://github.com/tamiespinosa/stock-analysis/blob/a36556cee6e784b0aa7973acf9afcac611f73115/VBA_Challenge.xlsm)

[2] https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax
