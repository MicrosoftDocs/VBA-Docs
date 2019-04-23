---
title: GroupShapes object (Word)
ms.prod: word
ms.assetid: de29d571-476b-fa8b-619e-f7d0181d9756
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupShapes object (Word)

Represents the individual shapes within a grouped shape. Each shape contained within a group of shapes is represented by a  **Shape** object.


## Remarks

Use the  **GroupItems** property to return the **GroupShapes** collection. Use **GroupItems** (Index), where Index is the number of the individual shape within the grouped shape, to return a single shape from the **GroupShapes** collection. The following example adds three triangles to the active document, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
With ActiveDocument.Shapes 
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



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]