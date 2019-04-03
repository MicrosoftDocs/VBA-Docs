---
title: ThreeDFormat object (Word)
keywords: vbawd10.chm2512
f1_keywords:
- vbawd10.chm2512
ms.prod: word
api_name:
- Word.ThreeDFormat
ms.assetid: d397e780-a53d-0cc3-7a02-b40397253e91
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat object (Word)

Represents a shape's three-dimensional formatting.


## Remarks

Use the  **ThreeD** property to return a **ThreeDFormat** object. The following example adds an oval to the active document and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Set myShape = ActiveDocument.Shapes _ 
 .AddShape(msoShapeOval, 90, 90, 90, 40) 
With myShape.ThreeD 
 .Visible = True 
 .Depth = 50 
 ' RGB value for purple 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With
```

You cannot apply three-dimensional formatting to some kinds of shapes, such as beveled shapes or multiple-disjoint paths. Most of the properties and methods of the  **ThreeDFormat** object for such a shape will fail.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]