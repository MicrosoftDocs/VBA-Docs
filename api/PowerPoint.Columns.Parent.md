---
title: Columns.Parent property (PowerPoint)
keywords: vbapp10.chm623002
f1_keywords:
- vbapp10.chm623002
ms.prod: powerpoint
api_name:
- PowerPoint.Columns.Parent
ms.assetid: 997abbd6-9a36-e80d-d1d8-543c182ca51a
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Columns](PowerPoint.Columns.md) object.


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


[Columns Object](PowerPoint.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]