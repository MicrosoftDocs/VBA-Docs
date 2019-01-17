---
title: PrintRanges.Parent Property (PowerPoint)
keywords: vbapp10.chm518005
f1_keywords:
- vbapp10.chm518005
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRanges.Parent
ms.assetid: 95bacc46-413d-2694-6ac2-7883609e26c7
ms.date: 06/08/2017
localization_priority: Normal
---


# PrintRanges.Parent Property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

 _expression_. `Parent`

 _expression_ A variable that represents a [PrintRanges](./PowerPoint.PrintRanges.md) object.


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


[PrintRanges Object](PowerPoint.PrintRanges.md)

