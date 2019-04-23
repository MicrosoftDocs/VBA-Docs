---
title: Area3DGroup property (Excel Graph)
keywords: vbagr10.chm65553
f1_keywords:
- vbagr10.chm65553
ms.prod: excel
api_name:
- Excel.Area3DGroup
ms.assetid: 4274ca69-7a9d-76d9-7a57-814b6246d94f
ms.date: 04/09/2019
localization_priority: Normal
---


# Area3DGroup property (Excel Graph)

Returns a **ChartGroup** object that represents the specified area chart group on a 3D chart. Read-only **ChartGroup** object.

## Syntax

_expression_.**Area3DGroup**

_expression_ Required. An expression that returns a **[ChartGroup](excel.chartgroup-graph-object.md)** object.


## Example

This example turns on drop lines for the 3D area chart group.

```vb
myChart.Area3DGroup.HasDropLines = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]