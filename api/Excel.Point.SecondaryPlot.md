---
title: Point.SecondaryPlot property (Excel)
keywords: vbaxl10.chm576098
f1_keywords:
- vbaxl10.chm576098
ms.prod: excel
api_name:
- Excel.Point.SecondaryPlot
ms.assetid: 1a12020a-bbd5-30b0-106a-589a44b45ca6
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.SecondaryPlot property (Excel)

**True** if the point is in the secondary section of either a Pie of Pie chart or a Bar of Pie chart. Applies only to points on Pie of Pie charts or Bar of Pie charts. Read/write **Boolean**.


## Syntax

_expression_.**SecondaryPlot**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Example

This example must be run on either a Pie of Pie chart or a Bar of Pie chart. The example moves point four to the secondary section of the chart.

```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .Points(4).SecondaryPlot = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]