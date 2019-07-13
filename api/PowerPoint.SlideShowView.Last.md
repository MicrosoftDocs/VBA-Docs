---
title: SlideShowView.Last method (PowerPoint)
keywords: vbapp10.chm513018
f1_keywords:
- vbapp10.chm513018
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.Last
ms.assetid: 1188d75f-9561-b92c-e2d1-9ceb03eae904
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.Last method (PowerPoint)

Sets the specified slide show view to display the last slide in the presentation.


## Syntax

_expression_. `Last`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Remarks

If you use the  **Last** method to switch from one slide to another during a slide show, when you return to the original slide, its animation picks up where it left off.


## Example

This example sets slide show window one to display the last slide in the presentation.


```vb
SlideShowWindows(1).View.Last
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]