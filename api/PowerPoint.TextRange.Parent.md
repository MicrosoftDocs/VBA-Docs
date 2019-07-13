---
title: TextRange.Parent property (PowerPoint)
keywords: vbapp10.chm569002
f1_keywords:
- vbapp10.chm569002
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Parent
ms.assetid: 303cc0cf-8c1c-60af-648e-fea4d25abb36
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


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


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]