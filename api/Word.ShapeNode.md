---
title: ShapeNode object (Word)
keywords: vbawd10.chm2509
f1_keywords:
- vbawd10.chm2509
ms.prod: word
api_name:
- Word.ShapeNode
ms.assetid: d5afb71a-a218-57f3-87f0-171094ba6610
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeNode object (Word)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The **ShapeNode** object is a member of the **ShapeNodes** collection. The **[ShapeNodes](Word.shapenodes.md)** collection contains all the nodes in a freeform.


## Remarks

Use  **Nodes** (Index), where Index is the node index number, to return a single **ShapeNode** object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.


```vb
With ActiveDocument.Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType 1, msoEditingSmooth 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]