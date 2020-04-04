---
title: ShapeRange.ThreeD property (PowerPoint)
keywords: vbapp10.chm548036
f1_keywords:
- vbapp10.chm548036
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ThreeD
ms.assetid: e0e2f72d-639b-86fd-2191-f537ddcd45ad
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.ThreeD property (PowerPoint)

Returns a **[ThreeDFormat](PowerPoint.ThreeDFormat.md)** object that contains 3D - effect formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**ThreeD**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

ThreeDFormat


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3D effects applied to shape one on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .Depth = 50

    'RGB value for purple

    .ExtrusionColor.RGB = RGB(255, 100, 255)

    .SetExtrusionDirection msoExtrusionTop

    .PresetLightingDirection = msoLightingLeft

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]