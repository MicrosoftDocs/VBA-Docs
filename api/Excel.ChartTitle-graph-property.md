---
title: ChartTitle property (Excel Graph)
keywords: vbagr10.chm5207199
f1_keywords:
- vbagr10.chm5207199
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: 736a91ad-a2ef-82c4-33b7-85c5ff78ae08
ms.date: 04/10/2019
localization_priority: Normal
---


# ChartTitle property (Excel Graph)

Returns a **ChartTitle** object that represents the title of the specified chart. Read-only.

## Syntax

_expression_.**ChartTitle**

_expression_ Required. An expression that returns a **[ChartTitle](Excel.ChartTitle-graph-object.md)** object.

## Example

This example sets the text for the title of the chart.

```vb
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]