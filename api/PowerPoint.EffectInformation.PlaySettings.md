---
title: EffectInformation.PlaySettings property (PowerPoint)
keywords: vbapp10.chm655008
f1_keywords:
- vbapp10.chm655008
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.PlaySettings
ms.assetid: 702cf5b9-8164-cd25-e441-566a9a94fc14
ms.date: 06/08/2017
localization_priority: Normal
---


# EffectInformation.PlaySettings property (PowerPoint)

Returns a  **[PlaySettings](PowerPoint.PlaySettings.md)** object that contains information about how the specified media clip plays during a slide show. Read-only.


## Syntax

_expression_. `PlaySettings`

_expression_ A variable that represents an [EffectInformation](PowerPoint.EffectInformation.md) object.


## Return value

PlaySettings


## Example

This example inserts a movie named Clock.avi onto slide one in the active presentation, sets it to play automatically after the slide transition, and specifies that the movie object be hidden during a slide show except when it is playing.


```vb
With ActivePresentation.Slides(1).Shapes.AddOLEObject(Left:=10, _
        Top:=10, Width:=250, Height:=250, _
    FileName:="c:\winnt\Clock.avi")
    With .AnimationSettings.PlaySettings
        .PlayOnEntry = True
        .HideWhileNotPlaying = True
    End With
End With
```


## See also


[EffectInformation Object](PowerPoint.EffectInformation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]