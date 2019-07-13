---
title: CommandEffect.Parent property (PowerPoint)
keywords: vbapp10.chm668002
f1_keywords:
- vbapp10.chm668002
ms.prod: powerpoint
api_name:
- PowerPoint.CommandEffect.Parent
ms.assetid: 06e0ba82-71ad-be99-adae-1ac44fd4d4bf
ms.date: 06/08/2017
localization_priority: Normal
---


# CommandEffect.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [CommandEffect](PowerPoint.CommandEffect.md) object.


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


[CommandEffect Object](PowerPoint.CommandEffect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]