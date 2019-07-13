---
title: Effect.Parent property (PowerPoint)
keywords: vbapp10.chm652002
f1_keywords:
- vbapp10.chm652002
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Parent
ms.assetid: 254fa25b-ef29-c2fe-313d-daadba3e8db4
ms.date: 06/08/2017
localization_priority: Normal
---


# Effect.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an [Effect](PowerPoint.Effect.md) object.


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



[Effect Object](PowerPoint.Effect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]