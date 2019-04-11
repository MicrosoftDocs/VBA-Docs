---
title: MinorGridlines property (Excel Graph)
keywords: vbagr10.chm5207695
f1_keywords:
- vbagr10.chm5207695
ms.prod: excel
api_name:
- Excel.MinorGridlines
ms.assetid: 80ca57a1-7e8f-4d83-0da6-2a28399c27af
ms.date: 04/11/2019
localization_priority: Normal
---


# MinorGridlines property (Excel Graph)

Returns a **Gridlines** object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Syntax

_expression_.**MinorGridlines**

_expression_ An expression that returns a **[Gridlines](Excel.Gridlines-graph-object.md)** object.

## Example

This example sets the color of the minor gridlines for the value axis in the chart to blue.

```vb
With myChart.Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 5 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]