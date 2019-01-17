---
title: SlideShowTransition.Parent Property (PowerPoint)
keywords: vbapp10.chm539002
f1_keywords:
- vbapp10.chm539002
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.Parent
ms.assetid: 32ab0ea5-ad24-ba48-6c00-31a1912c8d67
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowTransition.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [SlideShowTransition](./PowerPoint.SlideShowTransition.md) object.


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


[SlideShowTransition Object](PowerPoint.SlideShowTransition.md)

