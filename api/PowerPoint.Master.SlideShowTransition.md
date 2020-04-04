---
title: Master.SlideShowTransition property (PowerPoint)
keywords: vbapp10.chm533016
f1_keywords:
- vbapp10.chm533016
ms.prod: powerpoint
api_name:
- PowerPoint.Master.SlideShowTransition
ms.assetid: 935cadd9-a57a-a792-62b4-e198546438b2
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.SlideShowTransition property (PowerPoint)

Returns a **[SlideShowTransition](PowerPoint.SlideShowTransition.md)** object that represents the special effects for the specified slide transition. Read-only.


## Syntax

_expression_. `SlideShowTransition`

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


## Return value

SlideShowTransition


## Example

This example sets slide two in the active presentation to advance automatically after 5 seconds during a slide show and to play a dog bark sound at the slide transition.


```vb
With ActivePresentation.Slides(2).SlideShowTransition
    .AdvanceOnTime = True
    .AdvanceTime = 5
    .SoundEffect.ImportFromFile "c:\windows\media\dogbark.wav"
End With

ActivePresentation.SlideShowSettings.AdvanceMode = _
    ppSlideShowUseSlideTimings
```


## See also


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]