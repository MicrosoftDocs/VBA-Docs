---
title: Shape.GroupItems property (Word)
keywords: vbawd10.chm161480812
f1_keywords:
- vbawd10.chm161480812
ms.prod: word
api_name:
- Word.Shape.GroupItems
ms.assetid: c78ee480-b63a-cf0a-cbc0-94394f898912
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.GroupItems property (Word)

Returns a  **[GroupShapes](Word.groupshapes.md)** object that represents the individual shapes in the specified group. Read-only.


## Syntax

 _expression_. `GroupItems`

 _expression_ A variable that represents a '[Shape](Word.Shape.md)' object.


## Remarks

This property applies to  **Shape** object that represent grouped shapes. Use the **Item** method of the **[GroupShapes](Word.groupshapes.md)** object to return a single shape from the group.


## Example

This example adds three triangles to myDocument, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
Set myDocument = ActiveDocument 
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


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]