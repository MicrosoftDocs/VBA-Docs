---
title: Trendline.DisplayRSquared property (Excel)
keywords: vbaxl10.chm594080
f1_keywords:
- vbaxl10.chm594080
ms.prod: excel
api_name:
- Excel.Trendline.DisplayRSquared
ms.assetid: e8e447c3-d379-f6d0-74f2-629fa53b42ef
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.DisplayRSquared property (Excel)

**True** if the R-squared value of the trendline is displayed on the chart (in the same data label as the equation). Setting this property to **True** automatically turns on data labels. Read/write **Boolean**.


## Syntax

_expression_.**DisplayRSquared**

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