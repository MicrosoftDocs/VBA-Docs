---
title: ShapeNode Object (Publisher)
keywords: vbapb10.chm3604479
f1_keywords:
- vbapb10.chm3604479
ms.prod: publisher
api_name:
- Publisher.ShapeNode
ms.assetid: 8246e1fd-2477-91f4-490b-2d2b6032fccd
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeNode Object (Publisher)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The  **ShapeNode** object is a member of the **[ShapeNodes](Publisher.ShapeNodes.md)** collection. The **ShapeNodes** collection contains all the nodes in a freeform.
 


## Example

Use  **Nodes** (index), where index is the node index number, to return a single **ShapeNode** object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape one must be a freeform.
 

 

```vb
Sub ChangeNodeType() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType Index:=1, EditingType:=msoEditingSmooth 
 End If 
 End With 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Publisher.ShapeNode.Application.md)|
|[EditingType](Publisher.ShapeNode.EditingType.md)|
|[Parent](Publisher.ShapeNode.Parent.md)|
|[Points](Publisher.ShapeNode.Points.md)|
|[SegmentType](Publisher.ShapeNode.SegmentType.md)|

