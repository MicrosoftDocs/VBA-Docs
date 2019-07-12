---
title: SlideShowTransition.EntryEffect property (PowerPoint)
keywords: vbapp10.chm539006
f1_keywords:
- vbapp10.chm539006
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.EntryEffect
ms.assetid: 4a7bb737-a977-7a02-fccf-4bbb711a6375
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowTransition.EntryEffect property (PowerPoint)

Returns or sets the special effect applied to the specified slide transition. Read/write.


## Syntax

_expression_. `EntryEffect`

_expression_ A variable that represents a [SlideShowTransition](PowerPoint.SlideShowTransition.md) object.


## Return value

PpEntryEffect


## Remarks

If the  **[TextLevelEffect](PowerPoint.AnimationSettings.TextLevelEffect.md)** property for the specified shape is set to **ppAnimateLevelNone** (the default value) or the **[Animate](PowerPoint.AnimationSettings.Animate.md)** property is set to **False**, you won't see the special effect you've applied with the **EntryEffect** property.


## Example

This example adds a title slide to the active presentation and sets the title to fly in from the right whenever it is animated during a slide show.


```vb
With ActivePresentation.Slides.Add(1, ppLayoutTitleOnly).Shapes(1)

    .TextFrame.TextRange.Text = "Sample title"

    With .AnimationSettings

        .TextLevelEffect = ppAnimateByAllLevels

        .EntryEffect = ppEffectFlyFromRight

        .Animate = True

    End With

End With
```


## See also


[SlideShowTransition Object](PowerPoint.SlideShowTransition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]