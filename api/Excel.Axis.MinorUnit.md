---
title: Axis.MinorUnit property (Excel)
keywords: vbaxl10.chm561094
f1_keywords:
- vbaxl10.chm561094
ms.prod: excel
api_name:
- Excel.Axis.MinorUnit
ms.assetid: 64cd6523-19c3-7ebc-9b6b-db02667db4d2
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MinorUnit property (Excel)

Returns or sets the minor units on the value axis. Read/write **Double**.


## Syntax

_expression_.**MinorUnit**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting this property sets the **[MinorUnitIsAuto](Excel.Axis.MinorUnitIsAuto.md)** property to **False**.

Use the **[TickMarkSpacing](Excel.Axis.TickMarkSpacing.md)** property to set tick mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]