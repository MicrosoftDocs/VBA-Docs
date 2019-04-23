---
title: DisplayUnitCustom property (Excel Graph)
keywords: vbagr10.chm5241525
f1_keywords:
- vbagr10.chm5241525
ms.prod: excel
api_name:
- Excel.DisplayUnitCustom
ms.assetid: 18e2e0ae-13a9-3e45-6c93-90946ad98ebc
ms.date: 04/10/2019
localization_priority: Normal
---


# DisplayUnitCustom property (Excel Graph)

If the value returned or set by the **[DisplayUnit](Excel.DisplayUnit.md)** property is **xlCustom**, the **DisplayUnitCustom** property returns or sets the value of the units displayed for the value axis in the specified chart. The value must be a number from 0 through 10E307. Read/write **Double**.

## Syntax

_expression_.**DisplayUnitCustom**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Using unit labels for the value axis when charting large values makes the incremental labels on the axis more readable and the data easier to track. In other words, if you label your value axis in thousands (for example), you can use smaller numeric values next to the tick marks on the axis.


## Example

This example sets the units displayed on the value axis in _myChart_ to increments of 500.

```vb
With myChart.Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]