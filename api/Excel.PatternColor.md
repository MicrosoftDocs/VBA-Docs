---
title: PatternColor property
keywords: vbagr10.chm5207763
f1_keywords:
- vbagr10.chm5207763
ms.prod: excel
api_name:
- Excel.PatternColor
ms.assetid: f57dafd5-7690-67cd-013e-1cf31c26b570
ms.date: 06/08/2017
---


# PatternColor property

Returns or sets the color of the interior pattern as an RGB value. Read/write **Variant**.

_expression_. PatternColor

_expression_ Required. An expression that returns an [Interior](https://docs.microsoft.com/office/vba/api/Excel.Interior(Graph%20property)) object.


## Example

This example sets the color of the interior pattern for the chart area.

```vb
myChart.ChartArea.Interior.PatternColor = RGB(255,0,0)
```


