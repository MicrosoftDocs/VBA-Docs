---
title: AnimationSettings.SoundEffect property (PowerPoint)
keywords: vbapp10.chm565004
f1_keywords:
- vbapp10.chm565004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.SoundEffect
ms.assetid: b357a83d-167b-5429-7d7d-94851c8735ac
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.SoundEffect property (PowerPoint)

Returns a  **SoundEffect** object that represents the sound to be played during the transition to the specified slide. REad-only.


## Syntax

_expression_. `SoundEffect`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Return value

SoundEffect


## Example

This example sets the file Bass.wav to be played whenever shape one on slide one in the active presentation is animated.


```vb
With ActivePresentation.Slides(1).Shapes(1).AnimationSettings

    .Animate = True

    .TextLevelEffect = ppAnimateByAllLevels

    .SoundEffect.ImportFromFile "c:\bass.wav"

End With
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]