---
title: TabStops.Parent property (PowerPoint)
keywords: vbapp10.chm573002
f1_keywords:
- vbapp10.chm573002
ms.prod: powerpoint
api_name:
- PowerPoint.TabStops.Parent
ms.assetid: 5697b2b3-e2ad-343a-b52d-ab3b0bfd7ada
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStops.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [TabStops](PowerPoint.TabStops.md) object.


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


[TabStops Object](PowerPoint.TabStops.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]