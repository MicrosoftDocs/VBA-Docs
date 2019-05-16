---
title: TextFrame.Parent property (PowerPoint)
keywords: vbapp10.chm558001
f1_keywords:
- vbapp10.chm558001
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.Parent
ms.assetid: 3c5706c9-188e-6946-6e87-1501f32b1ce3
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


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


[TextFrame Object](PowerPoint.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]