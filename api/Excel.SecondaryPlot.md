---
title: SecondaryPlot property (Excel Graph)
keywords: vbagr10.chm5207958
f1_keywords:
- vbagr10.chm5207958
ms.prod: excel
api_name:
- Excel.SecondaryPlot
ms.assetid: 6806a9d3-06cc-3786-5d1e-fbc23680da7a
ms.date: 04/12/2019
localization_priority: Normal
---


# SecondaryPlot property (Excel Graph)

**True** if the point is in the secondary section of either a Pie of Pie chart or a Bar of Pie chart. Applies only to points on Pie of Pie charts or Bar of Pie charts. Read/write **Boolean**.


## Syntax

_expression_.**SecondaryPlot**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example must be run on either a Pie of Pie chart or a Bar of Pie chart. The example moves point four to the secondary section of the chart.

```vb
With myChart.SeriesCollection(1) 
 .Points(4).SecondaryPlot = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]