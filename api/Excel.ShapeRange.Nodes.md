---
title: ShapeRange.Nodes property (Excel)
keywords: vbaxl10.chm640111
f1_keywords:
- vbaxl10.chm640111
ms.prod: excel
api_name:
- Excel.ShapeRange.Nodes
ms.assetid: 6005d3f3-2c08-f539-87fc-51425ce81e0e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Nodes property (Excel)

Returns a  **[ShapeNodes](Excel.ShapeNodes.md)** collection that represents the geometric description of the specified shape.


## Syntax

_expression_. `Nodes`

_expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


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


[ShapeRange Object](Excel.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]