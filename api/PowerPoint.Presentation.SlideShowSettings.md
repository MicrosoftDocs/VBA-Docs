---
title: Presentation.SlideShowSettings property (PowerPoint)
keywords: vbapp10.chm583015
f1_keywords:
- vbapp10.chm583015
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SlideShowSettings
ms.assetid: 90a5a5cb-1f78-bbb2-8e4c-eb35aae13c90
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.SlideShowSettings property (PowerPoint)

Returns a **[SlideShowSettings](PowerPoint.SlideShowSettings.md)** object that represents the slide show settings for the specified presentation. Read-only.


## Syntax

_expression_. `SlideShowSettings`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

SlideShowSettings


## Example

This example starts a slide show meant to be presented by a speaker. The slide show will run with animation and narration turned off.


```vb
With Application.ActivePresentation.SlideShowSettings

    .ShowType = ppShowTypeSpeaker

    .ShowWithNarration = False

    .ShowWithAnimation = False

    .Run

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]