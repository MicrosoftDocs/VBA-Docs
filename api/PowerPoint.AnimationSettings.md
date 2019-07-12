---
title: AnimationSettings object (PowerPoint)
keywords: vbapp10.chm565000
f1_keywords:
- vbapp10.chm565000
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings
ms.assetid: ebbe4257-236b-35b4-bdf1-e92a1b4b417b
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings object (PowerPoint)

Represents the special effects applied to the animation for the specified shape during a slide show.


## Example

Use the [AnimationSettings](PowerPoint.Shape.AnimationSettings.md)property of the  **Shape** object to return the **AnimationSettings** object. The following example adds a slide that contains both a title and a three-item list to the active presentation, and then it sets the list to be animated by first-level paragraphs, to fly in from the left when animated, to dim to the specified color after being animated, and to animate its items in reverse order.


```vb
Set sObjs = ActivePresentation.Slides.Add(2, ppLayoutText).Shapes

sObjs.Title.TextFrame.TextRange.Text = "Top Three Reasons"

With sObjs.Placeholders(2)

    .TextFrame.TextRange.Text = _

        "Reason 1" & VBNewLine & "Reason 2" & VBNewLine & "Reason 3"

    With .AnimationSettings

        .TextLevelEffect = ppAnimateByFirstLevel

        .EntryEffect = ppEffectFlyFromLeft

        .AfterEffect = ppAfterEffectDim

        .DimColor.RGB = RGB(100, 120, 100)

        .AnimateTextInReverse = True

    End With

End With
```


## Properties



|Name|
|:-----|
|[AdvanceMode](PowerPoint.AnimationSettings.AdvanceMode.md)|
|[AdvanceTime](PowerPoint.AnimationSettings.AdvanceTime.md)|
|[AfterEffect](PowerPoint.AnimationSettings.AfterEffect.md)|
|[Animate](PowerPoint.AnimationSettings.Animate.md)|
|[AnimateBackground](PowerPoint.AnimationSettings.AnimateBackground.md)|
|[AnimateTextInReverse](PowerPoint.AnimationSettings.AnimateTextInReverse.md)|
|[AnimationOrder](PowerPoint.AnimationSettings.AnimationOrder.md)|
|[Application](PowerPoint.AnimationSettings.Application.md)|
|[ChartUnitEffect](PowerPoint.AnimationSettings.ChartUnitEffect.md)|
|[DimColor](PowerPoint.AnimationSettings.DimColor.md)|
|[EntryEffect](PowerPoint.AnimationSettings.EntryEffect.md)|
|[Parent](PowerPoint.AnimationSettings.Parent.md)|
|[PlaySettings](PowerPoint.AnimationSettings.PlaySettings.md)|
|[SoundEffect](PowerPoint.AnimationSettings.SoundEffect.md)|
|[TextLevelEffect](PowerPoint.AnimationSettings.TextLevelEffect.md)|
|[TextUnitEffect](PowerPoint.AnimationSettings.TextUnitEffect.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]