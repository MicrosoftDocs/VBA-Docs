---
title: Name property (Excel Graph)
keywords: vbagr10.chm3077554
f1_keywords:
- vbagr10.chm3077554
ms.prod: excel
ms.assetid: d3590902-6957-8e32-e627-5946ba66c44f
ms.date: 04/11/2019
localization_priority: Normal
---


# Name property (Excel Graph)

The **Name** property as it applies to the **Application**, **Trendline**, and **Font** objects and to all other objects.

## Application and Trendline objects

Returns or sets the name of the object. Read/write **String**.

### Syntax

_expression_.**Name**

_expression_ An expression that returns an **[Application](excel.application-graph-object.md)** or **[Trendline](excel.trendline-graph-object.md)** object.


### Example

This example assigns the name of the first trendline to the variable _myTrendname_.

```vb
myTrendname = myChart.SeriesCollection(1).Trendlines(1).Name
```

## Font object

Returns or sets the name of the object. Read/write **Variant**.

### Syntax

_expression_.**Name**

_expression_ Required. An expression that returns a **[Font](excel.font-graph-object.md)** object.



## All other objects

Returns or sets the name of the object. Read-only **String**.

### Syntax

_expression_.**Name**

_expression_ Required. An expression that returns one of the remaining objects.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]