---
title: AnimationBehaviors.Parent property (PowerPoint)
keywords: vbapp10.chm656002
f1_keywords:
- vbapp10.chm656002
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors.Parent
ms.assetid: 0202b2ae-d716-20f7-b984-240e58175523
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehaviors.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an [AnimationBehaviors](PowerPoint.AnimationBehaviors.md) object.


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


[AnimationBehaviors Object](PowerPoint.AnimationBehaviors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]