---
title: HasMinorGridlines property (Excel Graph)
keywords: vbagr10.chm65561
f1_keywords:
- vbagr10.chm65561
ms.prod: excel
api_name:
- Excel.HasMinorGridlines
ms.assetid: 78a690ee-0e5f-c69a-d2b3-54b2880f0933
ms.date: 04/11/2019
localization_priority: Normal
---


# HasMinorGridlines property (Excel Graph)

**True** if the axis has minor gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean**.

## Syntax

_expression_.**HasMinorGridlines**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the color of the minor gridlines for the value axis.

```vb
With myChart.Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 4 
 ' Set color to green. 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]