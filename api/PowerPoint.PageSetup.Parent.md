---
title: PageSetup.Parent property (PowerPoint)
keywords: vbapp10.chm527002
f1_keywords:
- vbapp10.chm527002
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.Parent
ms.assetid: 58f34a76-6ae8-c861-f4e7-0290931fb736
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [PageSetup](PowerPoint.PageSetup.md) object.


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


[PageSetup Object](PowerPoint.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]