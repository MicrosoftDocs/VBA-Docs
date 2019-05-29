---
title: Workbook.CreateForecastSheet method (Excel)
keywords: vbaxl10.chm199280
f1_keywords:
- vbaxl10.chm199280
ms.assetid: bec7b60b-7840-af15-6d5f-f5c184ea7aee
ms.date: 05/29/2019
ms.prod: excel
localization_priority: Normal
---


# Workbook.CreateForecastSheet method (Excel)

If you have historical time-based data, you can use **CreateForecastSheet** to create a forecast. When you create a forecast, a new worksheet is created that contains a table of the historical and predicted values and a chart showing this. A forecast can help you predict things like future sales, inventory requirements, or consumer trends.


## Syntax

_expression_.**CreateForecastSheet** (_Timeline_, _Values_, _ForecastStart_, _ForecastEnd_, _ConfInt_, _Seasonality_, _DataCompletion_, _Aggregation_, _ChartType_, _ShowStatsTable_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Parameter|Required/Optional|Data type|Description|
|:--------|:----------------|:--------|:----------|
| _Timeline_|Required|**Range**|The independent array or range of numeric data. The dates in the timeline must have a consistent step between them and can't be zero. The timeline isn't required to be sorted because the forecast mechanism will sort it implicitly for calculations. If a constant step can't be identified in the provided timeline, an invalid procedure call or argument (Error 5) will be returned.|
| _Values_|Required|**Range**| The historical values for which you want to forecast the next points.|
| _ForecastStart_|Optional|**Variant**|The point from which the generated forecast will begin.|
| _ForecastEnd_|Optional|**Variant**|The point in which the generated forecast will end.|
| _ConfInt_|Optional|**Variant**|A numerical value between 0 and 1 (exclusive), indicating a confidence level for the calculated confidence interval. For example, for a 90% confidence interval, a 90% confidence level will be computed (90% of future points are to fall within this radius from prediction). The default value is 95%.|
| _Seasonality_|Optional|**Variant**|A numerical value. The default value of 1 means Excel detects seasonality automatically for the forecast and uses positive, whole numbers for the length of the seasonal pattern. 0 indicates no seasonality, meaning the prediction will be linear. Positive whole numbers will indicate to the algorithm to use patterns of this length as the seasonality. For any other value, Error 5 will be returned. Maximum supported seasonality is 8,760 (number of hours in a year). Any seasonality above that number will result in the Error 5.|
| _DataCompletion_|Optional|**Variant**|Can be one of these **[XlForecastDataCompletion](Excel.xlforecastdatacompletion.md)** constants: **xlDataCompletionZeros** or **xlDataCompletionInterpolate** (default).|
| _Aggregation_|Optional|**Variant**|Can be one of these **[XlForecastAggregation](Excel.xlforecastaggregation.md)** constants: **xlAggregationAverage** (default), **xlAggregationCount**, **xlAggregationCountA**, **xlAggregationMax**, **xlAggregationMedian**, **xlAggregationMin**, or **xlAggregationSum**. |
| _ChartType_|Optional|**Variant**| Can be one of these **[XlForecastChartType](Excel.xlforecastcharttype.md)** constants: **xlChartTypeLine** (default) or **xlChartTypeColumn**. |
| _ShowStatsTable_|Optional|**Variant**| **True** or **False**. If **True**, an additional table is generated in the created sheet. This table contains statistical measures that indicate the accuracy of the created forecast.|

## Return value

None


## Remarks

When you use a formula to create a forecast, it returns a table with the historical and predicted data and a chart. The forecast predicts future values by using your existing time-based data and the AAA version of the Exponential Smoothing (ETS) algorithm. The table has the following columns, three of which are calculated columns:

- Historical time column (your time-based data series)
    
- Historical values column (your corresponding values data series)
    
- Forecasted values column (calculated by using FORECAST_ETS)
    
- Two columns representing the confidence interval (calculated by using FORECAST_ETS_CONFINT)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]