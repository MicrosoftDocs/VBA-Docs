---
title: ShapeRange.PickUp method (Excel)
keywords: vbaxl10.chm640087
f1_keywords:
- vbaxl10.chm640087
ms.prod: excel
api_name:
- Excel.ShapeRange.PickUp
ms.assetid: 6a7120d3-4fd4-cb4a-d838-89693267be22
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.PickUp method (Excel)

Copies the formatting of the specified shape. Use the  **[Apply](Excel.ShapeRange.Apply.md)** method to apply the copied formatting to another shape.


## Syntax

_expression_. `PickUp`

_expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


## Example

This example copies the formatting of shape one on  `myDocument` and then applies the copied formatting to shape two.


```vb
Set myDocument = Worksheets(1) 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

