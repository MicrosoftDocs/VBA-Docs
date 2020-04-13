---
title: SlideShowView.Slide property (PowerPoint)
keywords: vbapp10.chm513004
f1_keywords:
- vbapp10.chm513004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.Slide
ms.assetid: 4fdee96b-9b0d-64ba-19de-b810bf07987b
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.Slide property (PowerPoint)

Returns a **[Slide](PowerPoint.Slide.md)** object that represents the slide that's currently displayed in the specified slide show window view. Read-only.


## Syntax

_expression_. `Slide`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Return value

Slide


## Remarks

If the currently displayed slide is from an embedded presentation, you can use the  **[Parent](PowerPoint.Slide.Parent.md)** property of the **Slide** object returned by the **Slide** property to return the embedded presentation that contains the slide. (The **[Presentation](PowerPoint.SlideShowWindow.Presentation.md)** property of the **SlideShowWindow** object or **DocumentWindow** object returns the presentation from which the window was created, not the embedded presentation.)


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]