---
title: Bar3DGroup property (Excel Graph)
keywords: vbagr10.chm5207118
f1_keywords:
- vbagr10.chm5207118
ms.prod: excel
api_name:
- Excel.Bar3DGroup
ms.assetid: abbf8004-862f-7d70-2ca7-83f278d32836
ms.date: 04/09/2019
localization_priority: Normal
---


# Bar3DGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the specified bar chart group on a 3D chart. Read-only.

## Syntax

_expression_.**Bar3DGroup**

_expression_ Required. An expression that returns a **[ChartGroup](Excel.ChartGroup-graph-object.md)** object.

## Example

This example sets the space between bar clusters in the 3D bar chart group to be 50 percent of the bar width.

```vb
myChart.BarGroup3DGroup.GapWidth = 50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]