---
title: Shape.ActionSettings property (PowerPoint)
keywords: vbapp10.chm547048
f1_keywords:
- vbapp10.chm547048
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ActionSettings
ms.assetid: 67e76de6-c0c3-7a35-f01e-e1cab4eb13d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ActionSettings property (PowerPoint)

Returns an  **[ActionSettings](PowerPoint.ActionSettings.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.


## Syntax

_expression_. `ActionSettings`

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

ActionSettings


## Example

The following example sets the actions for clicking and moving the mouse over shape one on slide two in the active presentation.


```vb
Set myShape = ActivePresentation.Slides(2).Shapes(1)

myShape.ActionSettings(ppMouseClick).Action = ppActionLastSlide

myShape.ActionSettings(ppMouseOver).SoundEffect.Name = "applause"
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]