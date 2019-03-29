---
title: GroupShapes object (Excel)
keywords: vbaxl10.chm641072
f1_keywords:
- vbaxl10.chm641072
ms.prod: excel
api_name:
- Excel.GroupShapes
ms.assetid: 252d35da-9ab4-97f4-1e00-48ccfc003534
ms.date: 03/30/2019
localization_priority: Normal
---


# GroupShapes object (Excel)

Represents the individual shapes within a grouped shape.


## Remarks

Each shape is represented by a **[Shape](Excel.Shape.md)** object. Using the **[Item](Excel.Shapes.Item.md)** method with this object, you can work with single shapes within a group without having to ungroup them.


## Example

Use the **[GroupItems](Excel.Shape.GroupItems.md)** property of the **Shape** object to return the **GroupShapes** collection. 

Use **GroupItems** (_index_), where _index_ is the number of the individual shapes within the grouped shape, to return a single shape from the **GroupShapes** collection. 

The following example adds three triangles to _myDocument_, groups them, sets a color for the entire group, and then changes the color for the second triangle only.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 10, 10, 100, 100).Name = "shpOne" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 150, 10, 100, 100).Name = "shpTwo" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 300, 10, 100, 100).Name = "shpThree" 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble 
 End With 
End With
```

## Methods

- [Item](Excel.GroupShapes.Item.md)

## Properties

- [Application](Excel.GroupShapes.Application.md)
- [Count](Excel.GroupShapes.Count.md)
- [Creator](Excel.GroupShapes.Creator.md)
- [Parent](Excel.GroupShapes.Parent.md)
- [Range](Excel.GroupShapes.Range.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]