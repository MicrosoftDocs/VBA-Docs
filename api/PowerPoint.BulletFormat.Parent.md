---
title: BulletFormat.Parent Property (PowerPoint)
keywords: vbapp10.chm577002
f1_keywords:
- vbapp10.chm577002
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.Parent
ms.assetid: 95829267-e354-828b-5034-7da64dc5d5d7
ms.date: 06/08/2017
---


# BulletFormat.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [BulletFormat](./PowerPoint.BulletFormat.md) object.


### Return value

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


[BulletFormat Object](PowerPoint.BulletFormat.md)

