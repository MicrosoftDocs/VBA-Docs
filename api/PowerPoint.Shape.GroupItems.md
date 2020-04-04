---
title: Shape.GroupItems property (PowerPoint)
keywords: vbapp10.chm547023
f1_keywords:
- vbapp10.chm547023
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.GroupItems
ms.assetid: 295499de-0e74-e4ad-1145-f21927cbf2a9
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.GroupItems property (PowerPoint)

Returns a **[GroupShapes](PowerPoint.GroupShapes.md)** object that represents the individual shapes in the specified group. Use the **Item** method of the **GroupShapes** object to return a single shape from the group. Read-only.


## Syntax

_expression_. `GroupItems`

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

GroupShapes


## Example

This example adds three triangles to _myDocument_, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape(msoShapeIsoscelesTriangle, 10, _
        10, 100, 100).Name = "shpOne"

    .AddShape(msoShapeIsoscelesTriangle, 150, _
        10, 100, 100).Name = "shpTwo"

    .AddShape(msoShapeIsoscelesTriangle, 300, _
        10, 100, 100).Name = "shpThree"

    With .Range(Array("shpOne", "shpTwo", "shpThree")).Group
        .Fill.PresetTextured msoTextureBlueTissuePaper
        .GroupItems(2).Fill.PresetTextured msoTextureGreenMarble
    End With

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]