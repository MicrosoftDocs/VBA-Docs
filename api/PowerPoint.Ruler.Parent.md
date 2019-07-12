---
title: Ruler.Parent property (PowerPoint)
keywords: vbapp10.chm570002
f1_keywords:
- vbapp10.chm570002
ms.prod: powerpoint
api_name:
- PowerPoint.Ruler.Parent
ms.assetid: 2b3e5ae2-f303-53c8-5e5d-4d702f6127d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Ruler.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Ruler](PowerPoint.Ruler.md) object.


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


[Ruler Object](PowerPoint.Ruler.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]