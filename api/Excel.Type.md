---
title: Type Property
keywords: vbagr10.chm3077596
f1_keywords:
- vbagr10.chm3077596
ms.prod: excel
api_name:
- Excel.Type
ms.assetid: 467e47f2-3c6e-d52d-0fc7-26f3bca7c6f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Type Property

Type property as it applies to the  **Axis** object.

Returns or sets the axis type. Read/write XlAxisType .


|XlAxisType can be one of these XlAxisType constants.|
| **xlSeriesAxis**|
| **xlCategory**|
| **xlValue**|

_expression_. **Type**
 _expression_ Required. An expression that returns an [Axis](./Excel.Axis-graph-object.md) object.
Type property as it applies to the  **ChartColorFormat** object.
Returns the color type. Read-only Long.
_expression_. **Type**
 _expression_ Required. An expression that returns a [ChartColorFormat](./Excel.ChartColorFormat.md) object.
Type property as it applies to the  **ChartFillFormat** object.
Returns the fill type. Read-only MsoFillType .


|MsoFillType can be one of these MsoFillType constants.|
| **msoFillGradient**|
| **msoFillBackground**|
| **msoFillMixed**|
| **msoFillPatterned**|
| **msoFillPicture**|
| **msoFillSolid**|
| **msoFillTextured**|

_expression_. **Type**
 _expression_ Required. An expression that returns a [ChartFillFormat](./Excel.ChartFillFormat.md) object.
Type property as it applies to the  **DataLabel** and **DataLabels** objects.
Returns or sets the data label type. Read/write Variant.
_expression_. **Type**
 _expression_ Required. An expression that returns one of the above objects.
Type property as it applies to the  **Series** object.
Returns or sets the series type. Read/write Long.
_expression_. **Type**
 _expression_ Required. An expression that returns a [Series](./Excel.Series-graph-object.md) object.
Type property as it applies to the  **Trendline** object.
Returns or sets the trendline type. Read/write XlTrendlineType .


|XlTrendlineType can be one of these XlTrendlineType constants.|
| **xlExponential**|
| **xlLinear**|
| **xlLogarithmic**|
| **xlMovingAvg**|
| **xlPolynomial**|
| **xlPower**|

_expression_. **Type**
 _expression_ Required. An expression that returns a [Trendline](./Excel.Trendline-graph-object.md) object.

## Example

As it applies to the  **Trendline** object.

This example changes the trendline type for the first series in the chart. If the series has no trendline, this example fails.




```vb
myChart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]