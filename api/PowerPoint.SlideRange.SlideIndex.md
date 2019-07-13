---
title: SlideRange.SlideIndex property (PowerPoint)
keywords: vbapp10.chm532018
f1_keywords:
- vbapp10.chm532018
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.SlideIndex
ms.assetid: d913a70f-eb31-73b0-43bc-1021b3195a7e
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.SlideIndex property (PowerPoint)

Returns the index number of the specified slide within the  **Slides** collection. Read-only.


## Syntax

_expression_. `SlideIndex`

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Long


## Remarks

Unlike the  **SlideID** property, the **SlideIndex** property of a **Slide** object can change when you add slides to the presentation or rearrange the slides in the presentation. Therefore, using the **[FindBySlideID](PowerPoint.Slides.FindBySlideID.md)** method with the slide's ID number can be a more reliable way to return a specific **Slide** object from a **Slides** collection than using the **Item** method with the slide's index number.


## Example

This example displays the index number of the currently displayed slide in slide show window one.


```vb
MsgBox SlideShowWindows(1).View.Slide.SlideIndex
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]