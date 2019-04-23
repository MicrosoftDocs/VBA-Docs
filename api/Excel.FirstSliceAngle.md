---
title: FirstSliceAngle property (Excel Graph)
keywords: vbagr10.chm65599
f1_keywords:
- vbagr10.chm65599
ms.prod: excel
api_name:
- Excel.FirstSliceAngle
ms.assetid: 53f1fa5e-71d5-bf71-0fec-5f7be85b02d2
ms.date: 04/10/2019
localization_priority: Normal
---


# FirstSliceAngle property (Excel Graph)

Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3D pie, and doughnut charts. Read/write **Long**.

## Syntax

_expression_.**FirstSliceAngle**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example sets the angle for the first slice in chart group one. The example should be run on a 2D pie chart.

```vb
myChart.ChartGroups(1).FirstSliceAngle = 15
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]