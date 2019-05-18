---
title: Trendline.Forward2 property (Excel)
keywords: vbaxl10.chm594092
f1_keywords:
- vbaxl10.chm594092
ms.prod: excel
api_name:
- Excel.Trendline.Forward2
ms.assetid: af44bce5-8354-801e-f111-6adcb305b06b
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.Forward2 property (Excel)

Returns or sets the number of periods (or units on a scatter chart) that the trendline extends forward. Read/write **Double**.


## Syntax

_expression_.**Forward2**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Example

This example sets the number of units that the trendline on Chart1 extends forward and backward. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
With Charts("Chart1").SeriesCollection(1).Trendlines(1) 
 .Forward2 = 5 
 .Backward2 = .5 
End With 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]