---
title: Line3DGroup property (Excel Graph)
keywords: vbagr10.chm65556
f1_keywords:
- vbagr10.chm65556
ms.prod: excel
api_name:
- Excel.Line3DGroup
ms.assetid: 8624511a-41fe-52ea-2ea8-e45af74206c0
ms.date: 04/11/2019
localization_priority: Normal
---


# Line3DGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the line chart group on a 3D chart. Read-only.


## Syntax

_expression_.**Line3DGroup**

_expression_ An expression that returns a **[ChartGroup](Excel.ChartGroup-graph-object.md)** object .


## Example

This example sets the 3D line group to use a different color for each data marker.

```vb
myChart.Line3DGroup.VaryByCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]