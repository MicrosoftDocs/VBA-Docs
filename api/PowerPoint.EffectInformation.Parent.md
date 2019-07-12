---
title: EffectInformation.Parent property (PowerPoint)
keywords: vbapp10.chm655002
f1_keywords:
- vbapp10.chm655002
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.Parent
ms.assetid: 780fb3b7-8bdc-3b47-d5ce-b84c6b7c5b13
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectInformation.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an [EffectInformation](PowerPoint.EffectInformation.md) object.


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


[EffectInformation Object](PowerPoint.EffectInformation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]