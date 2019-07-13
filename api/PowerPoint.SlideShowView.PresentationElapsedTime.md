---
title: SlideShowView.PresentationElapsedTime property (PowerPoint)
keywords: vbapp10.chm513008
f1_keywords:
- vbapp10.chm513008
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.PresentationElapsedTime
ms.assetid: 6f710354-1691-4673-f83f-395d510d6999
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.PresentationElapsedTime property (PowerPoint)

Returns the number of seconds that have elapsed since the beginning of the specified slide show. Read-only.


## Syntax

_expression_. `PresentationElapsedTime`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Return value

Long


## Example

This example goes to slide seven in slide show window one if more than five minutes have elapsed since the beginning of the slide show.


```vb
With SlideShowWindows(1).View

    If .PresentationElapsedTime > 300 Then

        .GotoSlide 7

    End If

End With
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]