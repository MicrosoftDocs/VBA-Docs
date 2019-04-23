---
title: HasDisplayUnitLabel property (Excel Graph)
keywords: vbagr10.chm5241529
f1_keywords:
- vbagr10.chm5241529
ms.prod: excel
api_name:
- Excel.HasDisplayUnitLabel
ms.assetid: 5093286f-53ff-3c56-d047-7b6a92d2b7d6
ms.date: 04/11/2019
localization_priority: Normal
---


# HasDisplayUnitLabel property (Excel Graph)

**True** if the label specified by the **[DisplayUnit](Excel.DisplayUnit.md)** or **[DisplayUnitCustom](Excel.DisplayUnitCustom.md)** property is displayed on the value axis. **False** if no units are displayed. The default value is **True**. Read/write **Boolean**.

## Syntax

_expression_.**HasDisplayUnitLabel**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the units on the value axis in _myChart_ to increments of 500, but hides the unit label itself.

```vb
With myChart.Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .AxisTitle.Caption = "Rebate Amounts" 
 .HasDisplayUnitLabel = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]