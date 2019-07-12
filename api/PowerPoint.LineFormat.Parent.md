---
title: LineFormat.Parent property (PowerPoint)
keywords: vbapp10.chm553001
f1_keywords:
- vbapp10.chm553001
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.Parent
ms.assetid: 6644560e-0d3c-d675-b8a0-3481496c12ec
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [LineFormat](PowerPoint.LineFormat.md) object.


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


[LineFormat Object](PowerPoint.LineFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]