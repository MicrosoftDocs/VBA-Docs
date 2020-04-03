---
title: EffectInformation.AnimateBackground property (PowerPoint)
keywords: vbapp10.chm655004
f1_keywords:
- vbapp10.chm655004
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.AnimateBackground
ms.assetid: 37e9bfb5-3661-a3eb-d148-90d504f0e450
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectInformation.AnimateBackground property (PowerPoint)

Returns  **msoTrue** if the specified effect is a background animation. Read-only.


## Syntax

_expression_. `AnimateBackground`

_expression_ A variable that represents an [EffectInformation](PowerPoint.EffectInformation.md) object.


## Remarks

Use the [TextLevelEffect](PowerPoint.AnimationSettings.TextLevelEffect.md)and  **[TextUnitEffect](PowerPoint.EffectInformation.TextUnitEffect.md)** properties to control the animation of text attached to the specified shape.

If this property is set to  **msoTrue** and the **TextLevelEffect** property is set to **ppAnimateByAllLevels**, the shape and its text are animated simultaneously. If this property is set to **msoTrue** and the **TextLevelEffect** property is set to anything other than **ppAnimateByAllLevels**, the shape is animated immediately before the text is animated.

You won't see effects of setting this property unless the specified shape is animated. For a shape to be animated, the  **TextLevelEffect** property for the shape must be set to something other than **ppAnimateLevelNone**, and either the **[Animate](PowerPoint.AnimationSettings.Animate.md)** property must be set to **msoTrue**, or the **[EntryEffect](PowerPoint.AnimationSettings.EntryEffect.md)** property must be set to a constant other than **ppEffectNone**.

The value returned by the  **AnimateBackground** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified effect is not a background animation.|
|**msoTrue**| The specified effect is a background animation.|

## Example

This example changes the direction of the animation if the background is currently animated.


```vb
Sub ChangeAnimationDirection()

    With ActivePresentation.Slides(1).TimeLine.MainSequence(1)

        If .EffectInformation.AnimateBackground = msoTrue Then

            .EffectParameters.Direction = msoAnimDirectionTopLeft

        End If

    End With

End Sub
```


## See also


[EffectInformation Object](PowerPoint.EffectInformation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]