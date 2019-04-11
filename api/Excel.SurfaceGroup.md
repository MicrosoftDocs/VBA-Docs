---
title: SurfaceGroup property (Excel Graph)
keywords: vbagr10.chm65558
f1_keywords:
- vbagr10.chm65558
ms.prod: excel
api_name:
- Excel.SurfaceGroup
ms.assetid: f22bfac3-6c3c-0c82-8ca5-e167dd01e132
ms.date: 04/12/2019
localization_priority: Normal
---


# SurfaceGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the surface chart group of a 3D chart. Read-only **ChartGroup** object.

## Syntax

_expression_.**SurfaceGroup**

_expression_ Required. An expression that returns a **[ChartGroup](excel.chartgroup-graph-object.md)** object.


## Example

This example sets the 3D surface group to use a different color for each data marker. The example should be run on a 3D chart.

```vb
myChart.SurfaceGroup.VaryByCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]