---
title: Trendline.DisplayEquation property (Excel)
keywords: vbaxl10.chm594079
f1_keywords:
- vbaxl10.chm594079
ms.prod: excel
api_name:
- Excel.Trendline.DisplayEquation
ms.assetid: a9c3de54-5690-bf9b-505a-65b069195d53
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.DisplayEquation property (Excel)

**True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to **True** automatically turns on data labels. Read/write **Boolean**.


## Syntax

_expression_.**DisplayEquation**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Example

This example displays the R-squared value and equation for trendline one on Chart1. The example should be run on a 2D column chart that has a trendline for the first series.

```vb
With Charts("Chart1").SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]