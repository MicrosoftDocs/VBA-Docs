---
title: ShapeRange.AnimationSettings property (PowerPoint)
keywords: vbapp10.chm548047
f1_keywords:
- vbapp10.chm548047
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.AnimationSettings
ms.assetid: b248113c-54f6-5a36-b77f-63d76c10f7f3
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.AnimationSettings property (PowerPoint)

Returns an  **[AnimationSettings](PowerPoint.AnimationSettings.md)** object that represents all the special effects you can apply to the animation of the specified shape. Read-only.


## Syntax

_expression_. `AnimationSettings`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

AnimationSettings


## Example

This example sets shape one on slide two in the active presentation to fly in from the left when the slide is built.


```vb
With ActivePresentation.Slides(2).Shapes(1).AnimationSettings

    .EntryEffect = ppEffectFlyFromLeft

    .TextLevelEffect = ppAnimateByAllLevels

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]