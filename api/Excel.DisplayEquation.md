---
title: DisplayEquation property (Excel Graph)
keywords: vbagr10.chm5207312
f1_keywords:
- vbagr10.chm5207312
ms.prod: excel
api_name:
- Excel.DisplayEquation
ms.assetid: f3638bfd-d25d-96b4-5c20-2acf8703658d
ms.date: 04/10/2019
localization_priority: Normal
---


# DisplayEquation property (Excel Graph)

**True** if the equation for the trendline is displayed on the chart (in the same data label as the R-squared value). Setting this property to **True** automatically turns on data labels. Read/write **Boolean**.

## Syntax

_expression_.**DisplayEquation**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example displays the R-squared value and equation for trendline one. The example should be run on a 2D column chart that has a trendline for the first series.

```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = True 
 .DisplayEquation = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]