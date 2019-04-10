---
title: PlotArea property (Excel Graph)
keywords: vbagr10.chm65621
f1_keywords:
- vbagr10.chm65621
ms.prod: excel
api_name:
- Excel.PlotArea
ms.assetid: 047e8445-1197-2c9e-538d-5f77f6125c4c
ms.date: 04/11/2019
localization_priority: Normal
---


# PlotArea property (Excel Graph)

Returns a **PlotArea** object that represents the plot area of a chart. Read-only.


## Syntax

_expression_.**PlotArea**

_expression_ Required. An expression that returns a **[PlotArea](Excel.PlotArea-graph-object.md)** object.

## Example

This example sets the color of the plot area interior of _myChart_ to cyan.

```vb
myChart.PlotArea.Interior.ColorIndex = 8
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]