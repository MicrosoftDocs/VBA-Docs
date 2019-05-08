---
title: ChartGroup.FirstSliceAngle property (Excel)
keywords: vbaxl10.chm568077
f1_keywords:
- vbaxl10.chm568077
ms.prod: excel
api_name:
- Excel.ChartGroup.FirstSliceAngle
ms.assetid: a6bded62-d757-fc67-4677-7f9c12fd6395
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.FirstSliceAngle property (Excel)

Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3D pie, and doughnut charts. Can be a value from 0 through 360. Read/write **Long**.


## Syntax

_expression_.**FirstSliceAngle**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example sets the angle for the first slice in chart group one on Chart1. The example should be run on a 2D pie chart.

```vb
Charts("Chart1").ChartGroups(1).FirstSliceAngle = 15
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]