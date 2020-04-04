---
title: AnimationSettings.PlaySettings property (PowerPoint)
keywords: vbapp10.chm565010
f1_keywords:
- vbapp10.chm565010
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.PlaySettings
ms.assetid: 2cfd1ed9-7ed0-0f69-4df5-43aa22e37f46
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.PlaySettings property (PowerPoint)

Returns a **[PlaySettings](PowerPoint.PlaySettings.md)** object that contains information about how the specified media clip plays during a slide show. Read-only.


## Syntax

_expression_. `PlaySettings`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


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


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]