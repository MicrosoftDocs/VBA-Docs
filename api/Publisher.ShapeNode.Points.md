---
title: ShapeNode.Points property (Publisher)
keywords: vbapb10.chm3539201
f1_keywords:
- vbapb10.chm3539201
ms.prod: publisher
api_name:
- Publisher.ShapeNode.Points
ms.assetid: 30235d5a-9f05-4cc4-f62f-ac3cf4916e0d
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNode.Points property (Publisher)

Gets the _x-_ and _y-_ coordinates of the shape node. Read-only.


## Syntax

_expression_.**Points**

_expression_ A variable that represents a **[ShapeNode](Publisher.ShapeNode.md)** object.


## Remarks

This property is read-only. Use the **[SetPosition](Publisher.ShapeNodes.SetPosition.md)** method to set the location of the node.


## Example

This example moves node two in shape one on the first page of the active publication to the right 200 points and down 300 points. For this example to work, shape one must be a freeform drawing.

```vb
Sub SetPointsPosition() 
 Dim varArray As Variant 
 Dim intX As Integer 
 Dim intY As Integer 
 With ActiveDocument.Pages(1).Shapes(1).Nodes 
 varArray = .Item(2).Points 
 intX = varArray(1, 1) 
 intY = varArray(1, 2) 
 .SetPosition Index:=2, X1:=intX + 200, Y1:=intY + 300 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]