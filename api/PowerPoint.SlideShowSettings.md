---
title: SlideShowSettings object (PowerPoint)
keywords: vbapp10.chm514000
f1_keywords:
- vbapp10.chm514000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings
ms.assetid: d58c7c3b-a1cc-d819-b386-fd3fb7f967a2
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings object (PowerPoint)

Represents the slide show setup for a presentation.


## Example

Use the [SlideShowSettings](PowerPoint.Presentation.SlideShowSettings.md)property to return the  **SlideShowSettings** object. The first section in the following example sets all the slides in the active presentation to advance automatically after five seconds. The second section sets the slide show to start on slide two, end on slide four, advance slides by using the timings set in the first section, and run in a continuous loop until the user presses ESC. Finally, the example runs the slide show.


```vb
For Each s In ActivePresentation.Slides

    With s.SlideShowTransition

        .AdvanceOnTime = True

        .AdvanceTime = 5

    End With

Next



With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .AdvanceMode = ppSlideShowUseSlideTimings

    .LoopUntilStopped = True

    .Run

End With
```


## Methods



|Name|
|:-----|
|[Run](PowerPoint.SlideShowSettings.Run.md)|

## Properties



|Name|
|:-----|
|[AdvanceMode](PowerPoint.SlideShowSettings.AdvanceMode.md)|
|[Application](PowerPoint.SlideShowSettings.Application.md)|
|[EndingSlide](PowerPoint.SlideShowSettings.EndingSlide.md)|
|[LoopUntilStopped](PowerPoint.SlideShowSettings.LoopUntilStopped.md)|
|[NamedSlideShows](PowerPoint.SlideShowSettings.NamedSlideShows.md)|
|[Parent](PowerPoint.SlideShowSettings.Parent.md)|
|[PointerColor](PowerPoint.SlideShowSettings.PointerColor.md)|
|[RangeType](PowerPoint.SlideShowSettings.RangeType.md)|
|[ShowMediaControls](PowerPoint.SlideShowSettings.ShowMediaControls.md)|
|[ShowPresenterView](PowerPoint.SlideShowSettings.ShowPresenterView.md)|
|[ShowScrollbar](PowerPoint.SlideShowSettings.ShowScrollbar.md)|
|[ShowType](PowerPoint.SlideShowSettings.ShowType.md)|
|[ShowWithAnimation](PowerPoint.SlideShowSettings.ShowWithAnimation.md)|
|[ShowWithNarration](PowerPoint.SlideShowSettings.ShowWithNarration.md)|
|[SlideShowName](PowerPoint.SlideShowSettings.SlideShowName.md)|
|[StartingSlide](PowerPoint.SlideShowSettings.StartingSlide.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]