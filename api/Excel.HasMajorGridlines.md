---
title: HasMajorGridlines property (Excel Graph)
keywords: vbagr10.chm5207498
f1_keywords:
- vbagr10.chm5207498
ms.prod: excel
api_name:
- Excel.HasMajorGridlines
ms.assetid: f3c22d5d-4150-43b1-5f0d-3d49049e1e24
ms.date: 04/11/2019
localization_priority: Normal
---


# HasMajorGridlines property (Excel Graph)

**True** if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean**.

## Syntax

_expression_.**HasMajorGridlines**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the color of the major gridlines for the value axis.

```vb
With myChart.Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 3 'set color to red 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]