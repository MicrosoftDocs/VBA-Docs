---
title: Forward property (Excel Graph)
keywords: vbagr10.chm65727
f1_keywords:
- vbagr10.chm65727
ms.prod: excel
api_name:
- Excel.Forward
ms.assetid: 6a2e78d9-12ca-160a-7154-4968054f6b72
ms.date: 04/10/2019
localization_priority: Normal
---


# Forward property (Excel Graph)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write **Long**.

## Syntax

_expression_.**Forward**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the number of units that the trendline extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .Forward = 5 
 .Backward = .5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]