---
title: MinimumScaleIsAuto property (Excel Graph)
keywords: vbagr10.chm5207691
f1_keywords:
- vbagr10.chm5207691
ms.prod: excel
api_name:
- Excel.MinimumScaleIsAuto
ms.assetid: 95ed7a2b-efda-b05a-da2e-789a166a97c8
ms.date: 04/11/2019
localization_priority: Normal
---


# MinimumScaleIsAuto property (Excel Graph)

**True** if Graph calculates the minimum value for the axis. Read/write **Boolean**.

## Syntax

_expression_.**MinimumScaleIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting the **[MinimumScale](Excel.MinimumScale.md)** property sets this property to **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]