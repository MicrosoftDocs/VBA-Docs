---
title: SlideShowView.Next Method (PowerPoint)
keywords: vbapp10.chm513019
f1_keywords:
- vbapp10.chm513019
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.Next
ms.assetid: cf95eef7-4fd7-4c47-4436-037ec1882d4c
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.Next Method (PowerPoint)

Displays the slide immediately following the slide that's currently displayed. 


## Syntax

 _expression_. `Next`

 _expression_ A variable that represents a [SlideShowView](./PowerPoint.SlideShowView.md) object.


## Remarks

If the last slide is displayed, the  **Next** method closes the slide show in speaker mode and returns to the first slide in kiosk mode.

 Use the **[View](PowerPoint.SlideShowWindow.View.md)** property of the **SlideShowWindow** object to return the **SlideShowView** object.


## Example

This example shows the slide immediately following the currently displayed slide on slide show window one.


```vb
SlideShowWindows(1).View.Next
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]