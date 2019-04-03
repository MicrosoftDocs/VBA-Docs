---
title: FreeformBuilder object (PowerPoint)
keywords: vbapp10.chm546000
f1_keywords:
- vbapp10.chm546000
ms.prod: powerpoint
api_name:
- PowerPoint.FreeformBuilder
ms.assetid: fa188c8b-0781-dc9d-dd8d-3fc24c02d086
ms.date: 06/08/2017
localization_priority: Normal
---


# FreeformBuilder object (PowerPoint)

Represents the geometry of a freeform while it is being built.


## Example

Use the [BuildFreeform](PowerPoint.Shapes.BuildFreeform.md)method to return a  **FreeformBuilder** object. Use the [AddNodes](PowerPoint.FreeformBuilder.AddNodes.md)method to add nodes to the freefrom. Use the [ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)method to create the shape defined in the  **FreeformBuilder** object and add it to the **[Shapes](PowerPoint.Shapes.md)** collection. The following example adds a freeform with four segments to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    .AddNodes msoSegmentCurve, msoEditingCorner, _
        380, 230, 400, 250, 450, 300
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 40
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200
    .ConvertToShape
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]