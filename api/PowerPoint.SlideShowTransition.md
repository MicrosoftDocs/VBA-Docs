---
title: SlideShowTransition object (PowerPoint)
keywords: vbapp10.chm539000
f1_keywords:
- vbapp10.chm539000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition
ms.assetid: 60707d0d-62a8-0366-c22f-c5c5635fd762
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowTransition object (PowerPoint)

Contains information about how the specified slide advances during a slide show.


## Example

Use the [SlideShowTransition](PowerPoint.Slide.SlideShowTransition.md)property to return the  **SlideShowTransition** object. The following example specifies a Fast Strips Down-Left transition accompanied by the Bass.wav sound for slide one in the active presentation and specifies that the slide advance automatically five seconds after the previous animation or slide transition.


```vb
With ActivePresentation.Slides(1).SlideShowTransition

    .Speed = ppTransitionSpeedFast

    .EntryEffect = ppEffectStripsDownLeft

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

    .AdvanceOnTime = True

    .AdvanceTime = 5

End With

ActivePresentation.SlideShowSettings.AdvanceMode = _

    ppSlideShowUseSlideTimings
```


## Properties



|Name|
|:-----|
|[AdvanceOnClick](PowerPoint.SlideShowTransition.AdvanceOnClick.md)|
|[AdvanceOnTime](PowerPoint.SlideShowTransition.AdvanceOnTime.md)|
|[AdvanceTime](PowerPoint.SlideShowTransition.AdvanceTime.md)|
|[Application](PowerPoint.SlideShowTransition.Application.md)|
|[Duration](PowerPoint.SlideShowTransition.Duration.md)|
|[EntryEffect](PowerPoint.SlideShowTransition.EntryEffect.md)|
|[Hidden](PowerPoint.SlideShowTransition.Hidden.md)|
|[LoopSoundUntilNext](PowerPoint.SlideShowTransition.LoopSoundUntilNext.md)|
|[Parent](PowerPoint.SlideShowTransition.Parent.md)|
|[SoundEffect](PowerPoint.SlideShowTransition.SoundEffect.md)|
|[Speed](PowerPoint.SlideShowTransition.Speed.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]