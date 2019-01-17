---
title: Axis.DisplayUnitCustom property (Excel)
keywords: vbaxl10.chm561114
f1_keywords:
- vbaxl10.chm561114
ms.prod: excel
api_name:
- Excel.Axis.DisplayUnitCustom
ms.assetid: 77c660cc-dfb7-d4f7-6a8a-52522e026299
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.DisplayUnitCustom property (Excel)

If the value of the  **[DisplayUnit](Excel.Axis.DisplayUnit.md)** property is **xlCustom** , the **DisplayUnitCustom** property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write **Double**.


## Syntax

_expression_. `DisplayUnitCustom`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Remarks

Using unit labels when charting large values makes your tick mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example

This example sets the units displayed on the value axis in Chart1 to increments of 500.


```vb
With Charts("Chart1").Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]