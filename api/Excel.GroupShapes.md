---
title: GroupShapes Object (Excel)
keywords: vbaxl10.chm641072
f1_keywords:
- vbaxl10.chm641072
ms.prod: excel
api_name:
- Excel.GroupShapes
ms.assetid: 252d35da-9ab4-97f4-1e00-48ccfc003534
ms.date: 06/08/2017
---


# GroupShapes Object (Excel)

Represents the individual shapes within a grouped shape.


## Remarks

 Each shape is represented by a **[Shape](Excel.Shape.md)** object. Using the **[Item](Excel.Shapes.Item.md)** method with this object, you can work with single shapes within a group without having to ungroup them.


## Example

Use the  **[GroupItems](Excel.Shape.GroupItems.md)** property to return the **GroupShapes** collection. Use **[GroupItems](Excel.Shape.GroupItems.md)** ( _index_ ), where _index_ is the number of the individual shape within the grouped shape, to return a single shape from the **GroupShapes** collection. The following example adds three triangles to _myDocument_ , groups them, sets a color for the entire group, and then changes the color for the second triangle only.


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


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


