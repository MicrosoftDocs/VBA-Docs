---
title: Axis.MajorGridlines property (Excel)
keywords: vbaxl10.chm561084
f1_keywords:
- vbaxl10.chm561084
ms.prod: excel
api_name:
- Excel.Axis.MajorGridlines
ms.assetid: 618f880a-2b5d-2357-3c85-7b4858723b28
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorGridlines property (Excel)

Returns a  **[Gridlines](Excel.Gridlines(object).md)** object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Syntax

_expression_. `MajorGridlines`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Example

This example sets the color of the major gridlines for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMajorGridlines Then 
 .MajorGridlines.Border.ColorIndex = 5 'set color to blue 
 End If 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

