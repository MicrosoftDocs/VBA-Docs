---
title: SlideShowSettings.EndingSlide property (PowerPoint)
keywords: vbapp10.chm514006
f1_keywords:
- vbapp10.chm514006
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.EndingSlide
ms.assetid: 50489e3a-bdfe-b495-97d1-69ba1d7bf2b9
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.EndingSlide property (PowerPoint)

Returns or sets the last slide to be displayed in the specified slide show. Read/write.


## Syntax

_expression_. `EndingSlide`

_expression_ A variable that represents an [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


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