---
title: Backward property (Excel Graph)
keywords: vbagr10.chm65721
f1_keywords:
- vbagr10.chm65721
ms.prod: excel
api_name:
- Excel.Backward
ms.assetid: a92f33cb-45cd-baea-57e1-d76f44b041cb
ms.date: 04/09/2019
localization_priority: Normal
---


# Backward property (Excel Graph)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends backward. Read/write **Long**.

## Syntax

_expression_.**Backward**

_expression_ Required. An expression that returns a **[Trendline](excel.trendline-graph-object.md)** object.

## Example

This example sets the number of units that the trendline extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]