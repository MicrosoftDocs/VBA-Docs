---
title: Placeholders.Parent Property (PowerPoint)
keywords: vbapp10.chm544002
f1_keywords:
- vbapp10.chm544002
ms.prod: powerpoint
api_name:
- PowerPoint.Placeholders.Parent
ms.assetid: 216faab2-d0cc-1967-3b96-32bdea5a9b72
ms.date: 06/08/2017
localization_priority: Normal
---


# Placeholders.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [Placeholders](./PowerPoint.Placeholders.md) object.


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


[Placeholders Object](PowerPoint.Placeholders.md)

