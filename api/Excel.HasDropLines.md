---
title: HasDropLines property (Excel Graph)
keywords: vbagr10.chm5207474
f1_keywords:
- vbagr10.chm5207474
ms.prod: excel
api_name:
- Excel.HasDropLines
ms.assetid: 31f00864-86bc-9237-bf93-b52ab8cd1b59
ms.date: 04/11/2019
localization_priority: Normal
---


# HasDropLines property (Excel Graph)

**True** if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write **Boolean**.

## Syntax

_expression_.**HasDropLines**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

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