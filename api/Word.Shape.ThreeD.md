---
title: Shape.ThreeD property (Word)
keywords: vbawd10.chm161480826
f1_keywords:
- vbawd10.chm161480826
ms.prod: word
api_name:
- Word.Shape.ThreeD
ms.assetid: 35657b12-0967-5a54-6f12-b87119f51005
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ThreeD property (Word)

Returns a  **ThreeDFormat** object that contains 3D formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**ThreeD**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example sets the depth, extrusion color, extrusion direction, and lighting direction for the 3D effects applied to shape one on _myDocument_.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 50 
 ' RGB value for purple 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
 .SetExtrusionDirection msoExtrusionTop 
 .PresetLightingDirection = msoLightingLeft 
End With
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]