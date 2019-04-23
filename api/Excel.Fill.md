---
title: Fill property (Excel Graph)
keywords: vbagr10.chm5207366
f1_keywords:
- vbagr10.chm5207366
ms.prod: excel
api_name:
- Excel.Fill
ms.assetid: 7a8ea56d-1b39-cc70-1fbc-7d1a488b1aba
ms.date: 04/10/2019
localization_priority: Normal
---


# Fill property (Excel Graph)

Returns a **ChartFillFormat** object that contains fill formatting properties for the specified chart. Read-only.

## Syntax

_expression_.**Fill**

_expression_ Required. An expression that returns a **[ChartFillFormat](Excel.ChartFillFormat.md)** object.

## Example

This example sets the fill format for the chart to the preset brass color.

```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .PresetGradient msoGradientDiagonalDown, 3, msoGradientBrass 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]