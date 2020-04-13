---
title: TextRange.ActionSettings property (PowerPoint)
keywords: vbapp10.chm569003
f1_keywords:
- vbapp10.chm569003
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.ActionSettings
ms.assetid: 7a66ca28-d6b9-2066-4c88-a04888d5ba1e
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.ActionSettings property (PowerPoint)

Returns an **[ActionSettings](PowerPoint.ActionSettings.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.


## Syntax

_expression_. `ActionSettings`

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


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


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]