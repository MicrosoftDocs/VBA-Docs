---
title: SlideShowView.ResetSlideTime method (PowerPoint)
keywords: vbapp10.chm513024
f1_keywords:
- vbapp10.chm513024
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.ResetSlideTime
ms.assetid: aa00c585-d3c3-9cdc-860d-8c1f2f0a6ef3
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.ResetSlideTime method (PowerPoint)

Resets the elapsed time (represented by the  **[SlideElapsedTime](PowerPoint.SlideShowView.SlideElapsedTime.md)** property) for the slide that's currently displayed to 0 (zero).


## Syntax

_expression_. `ResetSlideTime`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Example

This example resets the elapsed time for the slide that's currently displayed in slide show window one to 0 (zero).


```vb
SlideShowWindows(1).View.ResetSlideTime
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]