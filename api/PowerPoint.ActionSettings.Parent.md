---
title: ActionSettings.Parent property (PowerPoint)
keywords: vbapp10.chm566002
f1_keywords:
- vbapp10.chm566002
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings.Parent
ms.assetid: d0c6c5db-5117-36af-5703-c79010903646
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSettings.Parent property (PowerPoint)

Returns the parent object for the specified object.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an **[ActionSettings](PowerPoint.ActionSettings.md)** object.


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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]