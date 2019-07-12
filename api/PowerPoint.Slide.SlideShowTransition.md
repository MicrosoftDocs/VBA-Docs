---
title: Slide.SlideShowTransition property (PowerPoint)
keywords: vbapp10.chm531005
f1_keywords:
- vbapp10.chm531005
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.SlideShowTransition
ms.assetid: bb931628-0ad1-e58b-9ddb-5680cb6ce9ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.SlideShowTransition property (PowerPoint)

Returns a  **[SlideShowTransition](PowerPoint.SlideShowTransition.md)** object that represents the special effects for the specified slide transition. Read-only.


## Syntax

_expression_. `SlideShowTransition`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


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


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]