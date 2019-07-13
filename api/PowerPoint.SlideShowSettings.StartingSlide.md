---
title: SlideShowSettings.StartingSlide property (PowerPoint)
keywords: vbapp10.chm514005
f1_keywords:
- vbapp10.chm514005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.StartingSlide
ms.assetid: e7afc69c-0224-b22a-fc23-bb985e710c1a
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.StartingSlide property (PowerPoint)

Returns or sets the first slide to be displayed in the specified slide show. Read/write.


## Syntax

_expression_. `StartingSlide`

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

Long


## Example

This example runs a slide show of the active presentation, starting with slide two and ending with slide four.


```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .Run

End With
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]