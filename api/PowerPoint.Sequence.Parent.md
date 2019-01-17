---
title: Sequence.Parent Property (PowerPoint)
keywords: vbapp10.chm651002
f1_keywords:
- vbapp10.chm651002
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.Parent
ms.assetid: fffc3d75-fd32-c27f-7c9f-b999d35e0ff3
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [Sequence](./PowerPoint.Sequence.md) object.


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


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]