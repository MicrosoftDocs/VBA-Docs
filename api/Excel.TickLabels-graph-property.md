---
title: TickLabels property (Excel Graph)
keywords: vbagr10.chm65627
f1_keywords:
- vbagr10.chm65627
ms.prod: excel
api_name:
- Excel.TickLabels
ms.assetid: 5aa48053-c9ff-71c7-7a03-d7fe47e681c7
ms.date: 04/12/2019
localization_priority: Normal
---


# TickLabels property (Excel Graph)

Returns a **TickLabels** object that represents the tick-mark labels for the specified axis. Read-only.

## Syntax

_expression_.**TickLabels**

_expression_ An expression that returns a **[TickLabels](Excel.TickLabels-graph-object.md)** object.

## Example

This example sets the color of the tick-mark label font for the value axis.

```vb
myChart.Axes(xlValue).TickLabels.Font.ColorIndex = 3
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]