---
title: GroupShapes object (PowerPoint)
keywords: vbapp10.chm549000
f1_keywords:
- vbapp10.chm549000
ms.prod: powerpoint
api_name:
- PowerPoint.GroupShapes
ms.assetid: db5eee43-c8bf-1d45-3075-8d9ceea2ae38
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupShapes object (PowerPoint)

Represents the individual shapes within a grouped shape. Each shape is represented by a  **[Shape](PowerPoint.Shape.md)** object. Using the [Item](PowerPoint.GroupShapes.Item.md)method with this object, you can work with single shapes within a group without having to ungroup them.


## Example

Use the [GroupItems](PowerPoint.Shape.GroupItems.md)property to return the  **GroupShapes** collection. Use **GroupItems** (_index_), where _index_ is the number of the individual shape within the grouped shape, to return a single shape from the **GroupShapes** collection. The following example adds three triangles to _myDocument_, groups them, sets a color for the entire group, and then changes the color for the second triangle only.


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


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]