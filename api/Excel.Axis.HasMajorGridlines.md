---
title: Axis.HasMajorGridlines property (Excel)
keywords: vbaxl10.chm561081
f1_keywords:
- vbaxl10.chm561081
ms.prod: excel
api_name:
- Excel.Axis.HasMajorGridlines
ms.assetid: 2cf9242a-79c5-8288-b71b-a5cd47d5abde
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.HasMajorGridlines property (Excel)

 **True** if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean**.


## Syntax

_expression_. `HasMajorGridlines`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Example

This example sets the color of the major gridlines for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 3 'set color to red 
 End If 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]