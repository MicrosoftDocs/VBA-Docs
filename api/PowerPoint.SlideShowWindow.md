---
title: SlideShowWindow object (PowerPoint)
keywords: vbapp10.chm507000
f1_keywords:
- vbapp10.chm507000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow
ms.assetid: 22468489-d4a2-ffea-7479-53ecb8d5da29
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowWindow object (PowerPoint)

Represents a window in which a slide show runs.


## Example

Use  **SlideShowWindows** (_index_), where _index_ is the slide show window index number, to return a single **SlideShowWindow** object. The following example activates slide show window two.


```vb
SlideShowWindows(2).Activate
```

Use the [Run](PowerPoint.SlideShowSettings.Run.md)method to create a new slide show window and return a reference to this slide show window. The following example runs a slide show of the active presentation and reduces the height of the slide show window just enough so that you can see the taskbar (for monitors with a screen resolution of 800 by 600).




```vb
With ActivePresentation.SlideShowSettings

    .ShowType = ppShowTypeSpeaker

    With .Run

        .Height = 300

        .Width = 400

    End With

End With
```

Use the [View](PowerPoint.SlideShowWindow.View.md)property to return the view in the specified slide show window. The following example sets the view in slide show window one to display slide three in the presentation.




```vb
SlideShowWindows(1).View.GotoSlide 3
```

Use the [Presentation](PowerPoint.SlideShowWindow.Presentation.md)property to return the presentation that's currently running in the specified slide show window. The following example displays the name of the presentation that's currently running in slide show window one.




```vb
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|Name|
|:-----|
|[DrawLine](PowerPoint.SlideShowView.DrawLine.md)|
|[EndNamedShow](PowerPoint.SlideShowView.EndNamedShow.md)|
|[EraseDrawing](PowerPoint.SlideShowView.EraseDrawing.md)|
|[Exit](PowerPoint.SlideShowView.Exit.md)|
|[First](PowerPoint.SlideShowView.First.md)|
|[FirstAnimationIsAutomatic](PowerPoint.SlideShowView.FirstAnimationIsAutomatic.md)|
|[GetClickCount](PowerPoint.SlideShowView.GetClickCount.md)|
|[GetClickIndex](PowerPoint.SlideShowView.GetClickIndex.md)|
|[GotoClick](PowerPoint.SlideShowView.GotoClick.md)|
|[GotoNamedShow](PowerPoint.SlideShowView.GotoNamedShow.md)|
|[GotoSlide](PowerPoint.SlideShowView.GotoSlide.md)|
|[Last](PowerPoint.SlideShowView.Last.md)|
|[Next](PowerPoint.SlideShowView.Next.md)|
|[Player](PowerPoint.SlideShowView.Player.md)|
|[Previous](PowerPoint.SlideShowView.Previous.md)|
|[ResetSlideTime](PowerPoint.SlideShowView.ResetSlideTime.md)|

## Properties



|Name|
|:-----|
|[AcceleratorsEnabled](PowerPoint.SlideShowView.AcceleratorsEnabled.md)|
|[AdvanceMode](PowerPoint.SlideShowView.AdvanceMode.md)|
|[Application](PowerPoint.SlideShowView.Application.md)|
|[CurrentShowPosition](PowerPoint.SlideShowView.CurrentShowPosition.md)|
|[IsNamedShow](PowerPoint.SlideShowView.IsNamedShow.md)|
|[LaserPointerEnabled](PowerPoint.slideshowview.laserpointerenabled.md)|
|[LastSlideViewed](PowerPoint.SlideShowView.LastSlideViewed.md)|
|[MediaControlsHeight](PowerPoint.SlideShowView.MediaControlsHeight.md)|
|[MediaControlsLeft](PowerPoint.SlideShowView.MediaControlsLeft.md)|
|[MediaControlsTop](PowerPoint.SlideShowView.MediaControlsTop.md)|
|[MediaControlsVisible](PowerPoint.SlideShowView.MediaControlsVisible.md)|
|[MediaControlsWidth](PowerPoint.SlideShowView.MediaControlsWidth.md)|
|[Parent](PowerPoint.SlideShowView.Parent.md)|
|[PointerColor](PowerPoint.SlideShowView.PointerColor.md)|
|[PointerType](PowerPoint.SlideShowView.PointerType.md)|
|[PresentationElapsedTime](PowerPoint.SlideShowView.PresentationElapsedTime.md)|
|[Slide](PowerPoint.SlideShowView.Slide.md)|
|[SlideElapsedTime](PowerPoint.SlideShowView.SlideElapsedTime.md)|
|[SlideShowName](PowerPoint.SlideShowView.SlideShowName.md)|
|[State](PowerPoint.SlideShowView.State.md)|
|[Zoom](PowerPoint.SlideShowView.Zoom.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]