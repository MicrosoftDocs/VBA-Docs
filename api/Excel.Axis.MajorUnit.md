---
title: Axis.MajorUnit property (Excel)
keywords: vbaxl10.chm561086
f1_keywords:
- vbaxl10.chm561086
ms.prod: excel
api_name:
- Excel.Axis.MajorUnit
ms.assetid: 6e58b341-6887-68c7-d0c1-a00abc226084
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorUnit property (Excel)

Returns or sets the major units for the value axis. Read/write  **Double**.


## Syntax

_expression_. `MajorUnit`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Remarks

Setting this property sets the  **[MajorUnitIsAuto](Excel.Axis.MajorUnitIsAuto.md)** property to **False**.

Use the  **[TickMarkSpacing](Excel.Axis.TickMarkSpacing.md)** property to set tick mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]