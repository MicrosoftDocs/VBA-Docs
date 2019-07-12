---
title: Slide.Parent property (PowerPoint)
keywords: vbapp10.chm531002
f1_keywords:
- vbapp10.chm531002
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Parent
ms.assetid: 02925312-0c0b-b1b9-c353-7d559f0e0050
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


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


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]