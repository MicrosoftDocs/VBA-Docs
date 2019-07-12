---
title: ColorSchemes.Parent property (PowerPoint)
keywords: vbapp10.chm536002
f1_keywords:
- vbapp10.chm536002
ms.prod: powerpoint
api_name:
- PowerPoint.ColorSchemes.Parent
ms.assetid: 5c59240a-c9a1-c6cc-ecc2-3e98dacd2a81
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorSchemes.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [ColorSchemes](PowerPoint.ColorSchemes.md) object.


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


[ColorSchemes Object](PowerPoint.ColorSchemes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]