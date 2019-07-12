---
title: AnimationSettings.Parent property (PowerPoint)
keywords: vbapp10.chm565002
f1_keywords:
- vbapp10.chm565002
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.Parent
ms.assetid: 73f01a7a-51c5-129f-34bf-2b7385e98ba5
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Return value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the  **Shape** object that contains the text.


```vb
Set myShapes = ActivePresentation.Slides(1).Shapes

With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _
        Top:=50, Width:=300, Height:=150).TextFrame
    .TextRange.Text = "Test text"
    .Parent.Rotation = 45
End With
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]