---
title: Border.ColorIndex property (Excel)
keywords: vbaxl10.chm547074
f1_keywords:
- vbaxl10.chm547074
ms.prod: excel
api_name:
- Excel.Border.ColorIndex
ms.assetid: 35e2dbbf-fd35-a08c-a969-bd08d0544d97
ms.date: 06/08/2017
localization_priority: Normal
---


# Border.ColorIndex property (Excel)

Returns or sets a  **Variant** value that represents the color of the border.


## Syntax

_expression_. `ColorIndex`

_expression_ A variable that represents a [Border](Excel.Border-graph-property.md) object.


## Remarks

The color is specified as an index value into the current color palette, or as one of the following  **[xlColorIndex](Excel.XlColorIndex.md)** constants:


-  **xlColorIndexAutomatic**
    
-  **xlColorIndexNone**
    

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


[Border Object](Excel.Border(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]