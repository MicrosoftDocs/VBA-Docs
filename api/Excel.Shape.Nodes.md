---
title: Shape.Nodes property (Excel)
keywords: vbaxl10.chm636104
f1_keywords:
- vbaxl10.chm636104
ms.prod: excel
api_name:
- Excel.Shape.Nodes
ms.assetid: 476b7ac6-d45c-c7a5-ef93-0cbe0c19ec15
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Nodes property (Excel)

Returns a  **[ShapeNodes](Excel.ShapeNodes.md)** collection that represents the geometric description of the specified shape.


## Syntax

_expression_. `Nodes`

_expression_ A variable that represents a [Shape](./Excel.Shape.md) object.


## Remarks

This property applies to  **[Shape](Excel.Shape.md)** or **[ShapeRange](Excel.ShapeRange.md)** objects that represent freeform drawings.


## Example

This example adds a smooth node with a curved segment after node four in shape three on  `myDocument`. Shape three must be a freeform drawing with at least four nodes.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100 
End With
```


## See also


[Shape Object](Excel.Shape.md)

