---
title: ShapeNode.EditingType property (Publisher)
keywords: vbapb10.chm3539200
f1_keywords:
- vbapb10.chm3539200
ms.prod: publisher
api_name:
- Publisher.ShapeNode.EditingType
ms.assetid: f01db634-b35a-48cd-851d-418848674686
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNode.EditingType property (Publisher)

If the specified node is a vertex, this property returns an **[MsoEditingType](Office.MsoEditingType.md)** constant indicating how changes made to the node affect the two segments connected to the node. If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. Read-only.


## Syntax

_expression_.**EditingType**

_expression_ A variable that represents a **[ShapeNode](Publisher.ShapeNode.md)** object.


## Return value

MsoEditingType


## Remarks

Use the **[SetEditingType](Publisher.ShapeNodes.SetEditingType.md)** method to set the value of this property.

The **EditingType** property value can be one of the **MsoEditingType** constants declared in the Microsoft Office type library.


## Example

This example changes all corner nodes to smooth curve nodes in the third shape in the active publication. The shape must be a freeform drawing.

```vb
Dim intNode As Integer 
 
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 For intNode = 1 to .Count 
 If .Item(intNode).EditingType = msoEditingCorner Then 
 .SetEditingType Index:=intNode, _ 
 EditingType:=msoEditingSmooth 
 End If 
 Next 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]