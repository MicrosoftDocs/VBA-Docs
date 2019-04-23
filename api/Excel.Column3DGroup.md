---
title: Column3DGroup property (Excel Graph)
keywords: vbagr10.chm3076976
f1_keywords:
- vbagr10.chm3076976
ms.prod: excel
api_name:
- Excel.Column3DGroup
ms.assetid: 9fa90f46-29b8-c710-93de-4150e276330c
ms.date: 04/10/2019
localization_priority: Normal
---


# Column3DGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the specified column chart group on a 3D chart. Read-only **ChartGroup** object.

## Syntax

_expression_.**Column3DGroup**

_expression_ Required. An expression that returns a **[ChartGroup](excel.chartgroup-graph-object.md)** object.


## Example

This example sets the space between column clusters in the 3D column chart group to be 50 percent of the column width.

```vb
myChart.Column3DGroup.GapWidth = 50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]