---
title: ShapeNode object (Excel)
keywords: vbaxl10.chm111000
f1_keywords:
- vbaxl10.chm111000
ms.prod: excel
api_name:
- Excel.ShapeNode
ms.assetid: c8b60d74-f11f-1659-30a3-6e180eb8bd58
ms.date: 04/02/2019
localization_priority: Normal
---


# ShapeNode object (Excel)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.


## Remarks

Nodes include the vertices between the segments of the freeform and the control points for curved segments. The **ShapeNode** object is a member of the **[ShapeNodes](Excel.ShapeNodes.md)** collection. The **ShapeNodes** collection contains all the nodes in a freeform.


## Example

Use **[Nodes](Excel.Shape.Nodes.md)** (_index_), where _index_ is the node index number, to return a single **ShapeNode** object. If node one in shape three on _myDocument_ is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType 1, msoEditingSmooth 
 End If 
End With
```

## Properties

- [Application](Excel.ShapeNode.Application.md)
- [Creator](Excel.ShapeNode.Creator.md)
- [EditingType](Excel.ShapeNode.EditingType.md)
- [Parent](Excel.ShapeNode.Parent.md)
- [Points](Excel.ShapeNode.Points.md)
- [SegmentType](Excel.ShapeNode.SegmentType.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]