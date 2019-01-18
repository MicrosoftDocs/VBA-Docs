---
title: GroupShapes Object (Publisher)
keywords: vbapb10.chm3407871
f1_keywords:
- vbapb10.chm3407871
ms.prod: publisher
api_name:
- Publisher.GroupShapes
ms.assetid: dd723f99-25a9-81cc-1d16-eb7dcd651c5e
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupShapes Object (Publisher)

Represents the individual shapes within a grouped shape. Each shape is represented by a  **[Shape](Publisher.Shape.md)** object. Using the **[Item](Publisher.GroupShapes.Item.md)** method with this object, you can work with single shapes within a group without having to ungroup them.
 


## Example

Use the  **[GroupItems](Publisher.Shape.GroupItems.md)** property to return a **GroupShapes** collection. Use **GroupItems** (index), where index is the number of the individual shape within the grouped shape, to return a single shape from the **GroupShapes** collection. The following example adds three triangles to the active document, groups them, sets a color for the entire group, and then changes the color for the third triangle only.
 

 

```vb
Sub WorkWithGroupShapes() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1).Shapes 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 50, 50, 100, 100).Name = "shpOne" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 200, 50, 100, 100).Name = "shpTwo" 
 .AddShape(msoShapeIsoscelesTriangle, _ 
 350, 50, 100, 100).Name = "shpThree" 
 With .Range(Array("shpOne", "shpTwo", "shpThree")).Group 
 .Fill.PresetTextured PresetTexture:=msoTextureBlueTissuePaper 
 .GroupItems(3).Fill.PresetTextured _ 
 PresetTexture:=msoTextureGreenMarble 
 End With 
 End With 
End Sub
```


## Methods



|Name|
|:-----|
|[Item](Publisher.GroupShapes.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.GroupShapes.Application.md)|
|[Count](Publisher.GroupShapes.Count.md)|
|[Parent](Publisher.GroupShapes.Parent.md)|

