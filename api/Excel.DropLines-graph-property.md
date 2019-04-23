---
title: DropLines property (Excel Graph)
keywords: vbagr10.chm5207331
f1_keywords:
- vbagr10.chm5207331
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 13dd4b80-669e-94c1-d592-439129d42d56
ms.date: 04/10/2019
localization_priority: Normal
---


# DropLines property (Excel Graph)

Returns a **DropLines** object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.

## Syntax

_expression_.**DropLines**

_expression_ Required. An expression that returns a **[DropLines](Excel.DropLines-graph-object.md)** object.

## Example

This example turns on drop lines for chart group one and then sets their line style, weight, and color. The example should be run on a 2D line chart that has one series.

```vb
With myChart.ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]