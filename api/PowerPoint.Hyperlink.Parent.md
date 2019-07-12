---
title: Hyperlink.Parent property (PowerPoint)
keywords: vbapp10.chm526002
f1_keywords:
- vbapp10.chm526002
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.Parent
ms.assetid: 5939f1a2-eb4f-d938-2683-530b0a408614
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Hyperlink](PowerPoint.Hyperlink.md) object.


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


[Hyperlink Object](PowerPoint.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]