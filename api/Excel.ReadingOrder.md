---
title: ReadingOrder property (Excel Graph)
keywords: vbagr10.chm5207913
f1_keywords:
- vbagr10.chm5207913
ms.prod: excel
api_name:
- Excel.ReadingOrder
ms.assetid: 70e3f325-0c75-24cb-d3e7-0273ce157061
ms.date: 04/12/2019
localization_priority: Normal
---


# ReadingOrder property (Excel Graph)

Returns or sets the reading order for the specified object. Can be one of the **[XlReadingOrder](word.xlreadingorder.md)** constants. Read/write **Long**.

## Syntax

_expression_.**ReadingOrder**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets the chart title's reading order to right-to-left.

```vb
myChart.ChartTitle.ReadingOrder = xlRTL
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]