---
title: PatternColor property (Excel Graph)
keywords: vbagr10.chm5207763
f1_keywords:
- vbagr10.chm5207763
ms.prod: excel
api_name:
- Excel.PatternColor
ms.assetid: f57dafd5-7690-67cd-013e-1cf31c26b570
ms.date: 04/11/2019
localization_priority: Normal
---


# PatternColor property (Excel Graph)

Returns or sets the color of the interior pattern as an RGB value. Read/write **Variant**.

## Syntax

_expression_.**PatternColor**

_expression_ Required. An expression that returns an **[Interior](excel.interior-graph-object.md)** object.


## Example

This example sets the color of the interior pattern for the chart area.

```vb
myChart.ChartArea.Interior.PatternColor = RGB(255,0,0)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]