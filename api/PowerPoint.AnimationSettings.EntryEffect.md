---
title: AnimationSettings.EntryEffect property (PowerPoint)
keywords: vbapp10.chm565005
f1_keywords:
- vbapp10.chm565005
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.EntryEffect
ms.assetid: de803113-6f7f-b1a2-1d52-43eeacccf666
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.EntryEffect property (PowerPoint)

For the  **AnimationSettings** object, this property returns or sets the special effect applied to the animation for the specified shape. Read/write.


## Syntax

_expression_. `EntryEffect`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


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


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]