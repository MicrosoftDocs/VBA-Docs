---
title: TickLabelSpacing property (Excel Graph)
keywords: vbagr10.chm5208063
f1_keywords:
- vbagr10.chm5208063
ms.prod: excel
api_name:
- Excel.TickLabelSpacing
ms.assetid: f8bf4611-3b25-3d66-f49b-5a088e95028b
ms.date: 04/12/2019
localization_priority: Normal
---


# TickLabelSpacing property (Excel Graph)

Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Read/write **Long**.

## Syntax

_expression_.**TickLabelSpacing**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Tick-mark label spacing on the value axis is always calculated by Graph.


## Example

This example sets the number of categories between tick-mark labels on the category axis.

```vb
myChart.Axes(xlCategory).TickLabelSpacing = 10
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]