---
title: Pie3DGroup property (Excel Graph)
keywords: vbagr10.chm5207791
f1_keywords:
- vbagr10.chm5207791
ms.prod: excel
api_name:
- Excel.Pie3DGroup
ms.assetid: 85e3866d-a38e-9749-c732-1e2d95a76c21
ms.date: 04/11/2019
localization_priority: Normal
---


# Pie3DGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the pie chart group on a 3D chart. Read-only.


## Syntax

_expression_.**Pie3DGroup**

_expression_ Required. An expression that returns a **[ChartGroup](Excel.ChartGroup-graph-object.md)** object.

## Example

This example sets the 3D pie group to use a different color for each data marker.

```vb
myChart.Pie3DGroup.VaryByCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]