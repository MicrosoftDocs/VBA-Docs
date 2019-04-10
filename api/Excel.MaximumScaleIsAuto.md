---
title: MaximumScaleIsAuto property (Excel Graph)
keywords: vbagr10.chm65572
f1_keywords:
- vbagr10.chm65572
ms.prod: excel
api_name:
- Excel.MaximumScaleIsAuto
ms.assetid: ca8115b8-0a45-0c88-5a5c-89c93d791452
ms.date: 04/11/2019
localization_priority: Normal
---


# MaximumScaleIsAuto property (Excel Graph)

**True** if Graph calculates the maximum value for the axis. Read/write **Boolean**.

## Syntax

_expression_.**MaximumScaleIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting the **[MaximumScale](Excel.MaximumScale.md)** property sets this property to **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]