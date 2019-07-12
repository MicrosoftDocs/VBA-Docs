---
title: EffectParameters.Parent property (PowerPoint)
keywords: vbapp10.chm654002
f1_keywords:
- vbapp10.chm654002
ms.prod: powerpoint
api_name:
- PowerPoint.EffectParameters.Parent
ms.assetid: b83fd852-e015-04f8-9856-ce018c23b848
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectParameters.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an [EffectParameters](PowerPoint.EffectParameters.md) object.


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


[EffectParameters Object](PowerPoint.EffectParameters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]