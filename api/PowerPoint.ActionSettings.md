---
title: ActionSettings object (PowerPoint)
keywords: vbapp10.chm566000
f1_keywords:
- vbapp10.chm566000
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings
ms.assetid: 8914c203-6b8d-fa80-16ad-7015595657b7
ms.date: 06/08/2017
localization_priority: Normal
---


# ActionSettings object (PowerPoint)

A collection that contains the two **[ActionSetting](PowerPoint.ActionSetting.md)** objects for a shape or text range. One **ActionSetting** object represents how the specified object reacts when the user clicks it during a slide show, and the other **ActionSetting** object represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.


## Example

Use the [ActionSettings](PowerPoint.Shape.ActionSettings.md)property to return the **ActionSettings** collection. Use **ActionSettings** (_index_), where _index_ is either **ppMouseClick** or **ppMouseOver**, to return a single **ActionSetting** object. The following example specifies that the CalculateTotal macro be run whenever the mouse pointer passes over the shape during a slide show.


```vb
With ActivePresentation.Slides(1).Shapes(3) _
        .ActionSettings(ppMouseOver)
    .Action = ppActionRunMacro
    .Run = "CalculateTotal"
    .AnimateAction = True
End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]