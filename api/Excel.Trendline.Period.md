---
title: Trendline.Period property (Excel)
keywords: vbaxl10.chm594088
f1_keywords:
- vbaxl10.chm594088
ms.prod: excel
api_name:
- Excel.Trendline.Period
ms.assetid: 142b675b-8859-a717-1e09-59a8b4000820
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.Period property (Excel)

Returns or sets the period for the moving-average trendline. Can be a value from 2 through 255. Read/write **Long**.


## Syntax

_expression_.**Period**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Example

This example sets the period for the moving-average trendline on Chart1. The example should be run on a 2D column chart with a single series that contains 10 data points and a moving-average trendline.

```vb
With Charts("Chart1").SeriesCollection(1).Trendlines(1) 
 If .Type = xlMovingAvg Then .Period = 5 
End With 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]