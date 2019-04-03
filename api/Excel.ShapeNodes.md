---
title: ShapeNodes object (Excel)
keywords: vbaxl10.chm112000
f1_keywords:
- vbaxl10.chm112000
ms.prod: excel
api_name:
- Excel.ShapeNodes
ms.assetid: 663721f1-8bd0-dd21-2362-fea2da3988bf
ms.date: 04/02/2019
localization_priority: Normal
---


# ShapeNodes object (Excel)

A collection of all the **[ShapeNode](Excel.ShapeNode.md)** objects in the specified freeform.


## Remarks

Each **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the **[BuildFreeform](Excel.Shapes.BuildFreeform.md)** and **[ConvertToShape](Excel.FreeformBuilder.ConvertToShape.md)** methods.


## Example

Use the **[Nodes](Excel.Shape.Nodes.md)** property of the **Shape** object to return the **ShapeNodes** collection. The following example deletes node four in shape three on _myDocument_. For this example to work, shape three must be a freeform with at least four nodes.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).Nodes.Delete 4
```

<br/>

Use the **Insert** method to create a new node and add it to the **ShapeNodes** collection. The following example adds a smooth node with a curved segment after node four in shape three on _myDocument_. For this example to work, shape three must be a freeform with at least four nodes.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100 
End With
```

<br/>

Use **Nodes** (_index_), where _index_ is the node index number, to return a single **ShapeNode** object. If node one in shape three on _myDocument_ is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType 1, msoEditingSmooth 
 End If 
End With
```

## Methods

- [Delete](Excel.ShapeNodes.Delete.md)
- [Insert](Excel.ShapeNodes.Insert.md)
- [Item](Excel.ShapeNodes.Item.md)
- [SetEditingType](Excel.ShapeNodes.SetEditingType.md)
- [SetPosition](Excel.ShapeNodes.SetPosition.md)
- [SetSegmentType](Excel.ShapeNodes.SetSegmentType.md)

## Properties

- [Application](Excel.ShapeNodes.Application.md)
- [Count](Excel.ShapeNodes.Count.md)
- [Creator](Excel.ShapeNodes.Creator.md)
- [Parent](Excel.ShapeNodes.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]