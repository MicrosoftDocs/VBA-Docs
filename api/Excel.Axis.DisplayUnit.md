---
title: Axis.DisplayUnit property (Excel)
keywords: vbaxl10.chm561113
f1_keywords:
- vbaxl10.chm561113
ms.prod: excel
api_name:
- Excel.Axis.DisplayUnit
ms.assetid: 81a4a639-aab4-e404-9e54-c75739cc57f9
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.DisplayUnit property (Excel)

Returns or sets the unit label for the value axis. Read/write **XlDisplayUnit**, **xlCustom**, or **xlNone**.


## Syntax

_expression_.**DisplayUnit**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

The unit label can be one of the **[XlDisplayUnit](Excel.XlDisplayUnit.md)** constants.

The unit label can also be one of the following constants: **xlCustom** or **xlNone**.

Using unit labels when charting large values makes your tick mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.


## Example

This example sets the units displayed on the value axis on Chart1 to hundreds.

```vb
With Charts("Chart1").Axes(xlValue) 
 .DisplayUnit = xlHundreds 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]