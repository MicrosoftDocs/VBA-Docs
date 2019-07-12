---
title: ActionSetting.Parent property (PowerPoint)
keywords: vbapp10.chm567002
f1_keywords:
- vbapp10.chm567002
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.Parent
ms.assetid: ade56ee1-5664-64a4-8936-1c80630a82fe
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSetting.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an **[ActionSetting](PowerPoint.ActionSetting.md)** object.


## Return value

Object


## Example

This example adds an oval containing text to slide one in the active presentation and rotates the oval and the text 45 degrees. The parent object for the text frame is the **Shape** object that contains the text.


```vb
Set myShapes = ActivePresentation.Slides(1).Shapes

With myShapes.AddShape(Type:=msoShapeOval, Left:=50, _
        Top:=50, Width:=300, Height:=150).TextFrame
    .TextRange.Text = "Test text"
    .Parent.Rotation = 45
End With
```


## See also


[ActionSetting Object](PowerPoint.ActionSetting.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]