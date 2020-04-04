---
title: SlideShowWindow.Presentation property (PowerPoint)
keywords: vbapp10.chm507004
f1_keywords:
- vbapp10.chm507004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow.Presentation
ms.assetid: 9c05deb7-a385-540f-97a5-1c5510f120c6
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowWindow.Presentation property (PowerPoint)

Returns a **[Presentation](PowerPoint.Presentation.md)** object that represents the presentation in which the specified document window or slide show window was created. Read-only.


## Syntax

_expression_. `Presentation`

_expression_ A variable that represents a [SlideShowWindow](PowerPoint.SlideShowWindow.md) object.


## Return value

Presentation


## Remarks

If the slide that's currently displayed in document window one is from an embedded presentation,  `Windows(1).View.Slide.Parent` returns the embedded presentation, and `Windows(1).Presentation` returns the presentation in which document window one was created.

If the slide that's currently displayed in slide show window one is from an embedded presentation,  `SlideShowWindows(1).View.Slide.Parent` returns the embedded presentation, and `SlideShowWindows(1).Presentation` returns the presentation in which the slide show was started.


## Example

This example continues the slide numbering for the presentation in window one into the slide numbering for the presentation in window two.


```vb
firstPresSlides = Windows(1).Presentation.Slides.Count
Windows(2).Presentation.PageSetup _
    .FirstSlideNumber = firstPresSlides + 1
```


## See also


[SlideShowWindow Object](PowerPoint.SlideShowWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]