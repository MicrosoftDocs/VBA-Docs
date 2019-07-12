---
title: SlideShowTransition.AdvanceTime property (PowerPoint)
keywords: vbapp10.chm539005
f1_keywords:
- vbapp10.chm539005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.AdvanceTime
ms.assetid: 79a120d2-5777-5eaa-a522-36e7d3bd539a
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowTransition.AdvanceTime property (PowerPoint)

Returns or sets the amount of time, in seconds, after which the specified slide transition will occur. Read/write.


## Syntax

_expression_. `AdvanceTime`

_expression_ A variable that represents a [SlideShowTransition](PowerPoint.SlideShowTransition.md) object.


## Return value

Single


## Remarks

To put the slide interval settings into effect for the entire slide show, set the  **[AdvanceMode](PowerPoint.SlideShowSettings.AdvanceMode.md)** property of the **[SlideShowSettings](PowerPoint.SlideShowSettings.md)** object to **ppSlideShowUseSlideTimings**.


## Example

This example sets slide one in the active presentation to advance after five seconds have passed or when the mouse is clicked — whichever occurs first.


```vb
With ActivePresentation.Slides(1).SlideShowTransition

    .AdvanceOnClick = msoTrue

    .AdvanceOnTime = msoTrue

    .AdvanceTime = 5

End With


```


## See also


[SlideShowTransition Object](PowerPoint.SlideShowTransition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]