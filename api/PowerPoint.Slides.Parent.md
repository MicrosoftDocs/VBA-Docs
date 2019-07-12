---
title: Slides.Parent property (PowerPoint)
keywords: vbapp10.chm530002
f1_keywords:
- vbapp10.chm530002
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.Parent
ms.assetid: 716b34cc-7e6f-e353-fa2e-82687f9d8784
ms.date: 06/08/2017
localization_priority: Normal
---


# Slides.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Slides](PowerPoint.Slides.md) object.


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


[Slides Object](PowerPoint.Slides.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]