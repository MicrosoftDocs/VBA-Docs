---
title: Type property (Excel Graph)
keywords: vbagr10.chm3077596
f1_keywords:
- vbagr10.chm3077596
ms.prod: excel
api_name:
- Excel.Type
ms.assetid: 467e47f2-3c6e-d52d-0fc7-26f3bca7c6f2
ms.date: 04/12/2019
localization_priority: Normal
---


# Type property (Excel Graph)

The **Type** property as it applies to the following objects.

## Axis object

Returns or sets the axis type. Read/write **[XlAxisType](excel.xlaxistype.md)**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns an **[Axis](Excel.Axis-graph-object.md)** object.


## ChartColorFormat object

Returns the color type. Read-only **Long**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns a **[ChartColorFormat](Excel.ChartColorFormat.md)** object.


## ChartFillFormat object

Returns the fill type. Read-only **[MsoFillType](office.msofilltype.md)**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns a **[ChartFillFormat](Excel.ChartFillFormat.md)** object.


## DataLabel and DataLabels objects

Returns or sets the data label type. Read/write **Variant**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns a **[DataLabel](excel.datalabel-graph-object.md)** object or **[DataLabels](excel.datalabels(collection).md)** collection.

## Series object

Returns or sets the series type. Read/write **Long**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns a **[Series](Excel.Series-graph-object.md)** object.

## Trendline object

Returns or sets the trendline type. Read/write **[XlTrendlineType](excel.xltrendlinetype.md)**.

### Syntax

_expression_.**Type**

_expression_ Required. An expression that returns a **[Trendline](Excel.Trendline-graph-object.md)** object.

### Example

This example changes the trendline type for the first series in the chart. If the series has no trendline, this example fails.

```vb
myChart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]