---
title: ShapeRange.GroupItems property (Excel)
keywords: vbaxl10.chm640104
f1_keywords:
- vbaxl10.chm640104
ms.prod: excel
api_name:
- Excel.ShapeRange.GroupItems
ms.assetid: daf6d12c-409a-cf0a-989f-319333d24596
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.GroupItems property (Excel)

Returns a **[GroupShapes](Excel.GroupShapes.md)** object that represents the individual shapes in the specified group. Use the **[Item](Excel.GroupShapes.Item.md)** method of the **GroupShapes** object to return a single shape from the group. Applies to **ShapeRange** objects that represent grouped shapes. Read-only.


## Syntax

_expression_.**GroupItems**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example adds three triangles to _myDocument_, groups them, sets a color for the entire group, and then changes the color for the second triangle only.

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]