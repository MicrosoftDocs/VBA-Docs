---
title: EffectInformation object (PowerPoint)
keywords: vbapp10.chm655000
f1_keywords:
- vbapp10.chm655000
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation
ms.assetid: 9b3d09f4-229b-8392-f9a4-777bf6557632
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectInformation object (PowerPoint)

Represents various animation options for an  **[Effect](PowerPoint.Effect.md)** object.


## Remarks

Use the members of the  **EffectInformation** object to return the current state of an **Effect** object, such as the after effect, whether the background animates along with its corresponding text, whether text animates in reverse, play settings, sound effects, text building behavior. All of the members of the **EffectInformation** object are read-only. To change any effect information properties, you must use the methods of the corresponding **[Sequence](PowerPoint.Sequence.md)** object.

Use the [EffectInformation](PowerPoint.Effect.EffectInformation.md)property of the  **Effect** object to return an **EffectInformation** object.


## Example

The following example sets the [HideWhileNotPlaying](PowerPoint.PlaySettings.HideWhileNotPlaying.md)property for the play settings in the main animation sequence.


```vb
Sub HideEffect()
    ActiveWindow.Selection.SlideRange(1).TimeLine _
        .MainSequence(1).EffectInformation.PlaySettings _
        .HideWhileNotPlaying = msoTrue
End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]