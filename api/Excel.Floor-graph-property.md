---
title: Floor property (Excel Graph)
keywords: vbagr10.chm65619
f1_keywords:
- vbagr10.chm65619
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: 1c82553e-7285-c759-416d-4537efd1c9ec
ms.date: 04/10/2019
localization_priority: Normal
---


# Floor property (Excel Graph)

Returns a **Floor** object that represents the floor of the 3D chart. Read-only.

## Syntax

_expression_.**Floor**

_expression_ Required. An expression that returns a **[Floor](Excel.Floor-graph-object.md)** object.


## Example

This example sets the floor color to blue. The example should be run on a 3D chart (the **Floor** property fails on 2D charts).

```vb
myChart.Floor.Interior.ColorIndex = 5
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]