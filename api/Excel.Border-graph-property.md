---
title: Border property (Excel Graph)
keywords: vbagr10.chm3076965
f1_keywords:
- vbagr10.chm3076965
ms.prod: excel
api_name:
- Excel.Border
ms.assetid: c4c01534-3d56-7496-0368-fea8d2e2d0ae
ms.date: 04/09/2019
localization_priority: Normal
---


# Border property (Excel Graph)

Returns a **Border** object that represents the border of the specified object. Read-only **Border** object.

## Syntax

_expression_.**Border**

_expression_ Required. An expression that returns a **[Border](Excel.Border-graph-object.md)** object.


## Example

This example sets the color of the chart area border to red.

```vb
myChart.ChartArea.Border.ColorIndex = 3
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
