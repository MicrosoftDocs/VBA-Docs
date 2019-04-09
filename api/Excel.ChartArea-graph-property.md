---
title: ChartArea property (Excel Graph)
keywords: vbagr10.chm5207181
f1_keywords:
- vbagr10.chm5207181
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 1af59d11-2b63-d629-5dae-d9b9d8303ddf
ms.date: 04/10/2019
localization_priority: Normal
---


# ChartArea property (Excel Graph)

Returns a **ChartArea** object that represents the complete chart area for the chart. Read-only.

## Syntax

_expression_.**ChartArea**

_expression_ Required. An expression that returns a **[ChartArea](Excel.ChartArea-graph-object.md)** object.

## Example

This example sets the chart area interior color of _myChart_ to red, and sets the border color to blue.

```vb
With myChart.ChartArea 
    .Interior.ColorIndex = 3 
    .Border.ColorIndex = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]