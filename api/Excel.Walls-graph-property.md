---
title: Walls property (Excel Graph)
keywords: vbagr10.chm65622
f1_keywords:
- vbagr10.chm65622
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 74da4bfa-7b53-80d9-a673-42a67ffab787
ms.date: 04/12/2019
localization_priority: Normal
---


# Walls property (Excel Graph)

Returns a **Walls** object that represents the walls of the 3D chart. Read-only.


## Syntax

_expression_.**Walls**

_expression_ Required. An expression that returns a **[Walls](Excel.Walls-graph-object.md)** object.

## Remarks

This property doesn't apply to 3D pie charts.


## Example

This example sets the color of the wall border of the chart to red. The example should be run on a 3D chart.

```vb
myChart.Walls.Border.ColorIndex = 3
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]