---
title: TickLabelPosition property (Excel Graph)
keywords: vbagr10.chm3077594
f1_keywords:
- vbagr10.chm3077594
ms.prod: excel
api_name:
- Excel.TickLabelPosition
ms.assetid: 5b4b6bbc-5c0b-2428-b100-d3f3562d6927
ms.date: 04/12/2019
localization_priority: Normal
---


# TickLabelPosition property (Excel Graph)

Describes the position of tick-mark labels on the specified axis. Read/write **[XlTickLabelPosition](excel.xlticklabelposition.md)**.

## Syntax

_expression_.**TickLabelPosition**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets tick-mark labels on the category axis to the high position (above the chart).

```vb
myChart.Axes(xlCategory) _ 
 .TickLabelPosition = xlTickLabelPositionHigh
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]