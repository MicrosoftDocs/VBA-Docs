---
title: Add method (Excel Graph)
keywords: vbagr10.chm3077604
f1_keywords:
- vbagr10.chm3077604
ms.prod: excel
api_name:
- Excel.Add
ms.assetid: 529bbd0e-c726-2e88-fa75-d492fede7f37
ms.date: 04/06/2019
localization_priority: Normal
---


# Add method (Excel Graph)

Creates a new trendline. Returns a **Trendline** object.

## Syntax

_expression_.**Add** (_Type_, _Order_, _Period_, _Forward_, _Backward_, _Intercept_, _DisplayEquation_, _DisplayRSquared_, _Name_)

_expression_ Required. An expression that returns a **[Trendline](excel.trendline-graph-object.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_ | Optional |**[XlTrendlineType](excel.xltrendlinetype.md)** |The type of trendline. Can be one of the **XlTrendlineType** constants. |
|_Order_ |Optional |**Variant** |Required if _Type_ is **xlPolynomial**. The trendline order. Must be an integer from 2 through 6.|
|_Period_ |Optional |**Variant** |Required if _Type_ is **xlMovingAvg**. The trendline period. Must be an integer greater than 1 and less than the number of data points in the series you are adding a trendline to.|
|_Forward_ |Optional |**Variant**|The number of periods (or units on a scatter chart) that the trendline extends forward.|
|_Backward_ |Optional |**Variant**| The number of periods (or units on a scatter chart) that the trendline extends backward.|
|_Intercept_ |Optional |**Variant**|The trendline intercept. If this argument is omitted, the intercept is automatically set by the regression.|
|_DisplayEquation_ |Optional |**Variant**|**True** to display the equation of the trendline on the chart (in the same data label as the R-squared value). The default value is **False**.|
| _DisplayRSquared_| Optional |**Variant**|**True** to display the R-squared value of the trendline on the chart (in the same data label as the equation). The default value is **False**.|
| _Name_ |Optional |**Variant**|The name of the trendline, as text. If this argument is omitted, Graph generates a name.|

## Example

This example creates a new linear trendline on the chart.

```vb
myChart.SeriesCollection(1).Trendlines.Add
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]