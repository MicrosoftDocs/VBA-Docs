---
title: ShapeNode.Points property (Excel)
keywords: vbaxl10.chm111004
f1_keywords:
- vbaxl10.chm111004
ms.prod: excel
api_name:
- Excel.ShapeNode.Points
ms.assetid: fe09c78f-44c9-4e66-df7b-c23720216ec5
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNode.Points property (Excel)

Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in [points](../language/glossary/vbe-glossary.md#point). Read-only **Variant**.


## Syntax

_expression_.**Points**

_expression_ An expression that returns a **[ShapeNode](Excel.ShapeNode.md)** object.


## Return value

Variant


## Remarks

This property is read-only. Use the **[SetPosition](Excel.ShapeNodes.SetPosition.md)** method to set the value of this property.


## Example

This example moves node two in shape three on _myDocument_ to the right 200 points and down 300 points. Shape three must be a freeform drawing.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 pointsArray = .Item(2).Points 
 currXvalue = pointsArray(1, 1) 
 currYvalue = pointsArray(1, 2) 
 .SetPosition 2, currXvalue + 200, currYvalue + 300 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]