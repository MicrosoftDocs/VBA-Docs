---
title: Slide.SlideID property (PowerPoint)
keywords: vbapp10.chm531009
f1_keywords:
- vbapp10.chm531009
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.SlideID
ms.assetid: 9d2d920c-a876-c71c-083f-ae8a3ad06c85
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.SlideID property (PowerPoint)

Returns a unique ID number for the specified slide. Read-only.


## Syntax

_expression_. `SlideID`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

Long


## Remarks

Unlike the  **SlideIndex** property, the **SlideID** property of a **Slide** object won't change when you add slides to the presentation or rearrange the slides in the presentation. Therefore, using the **[FindBySlideID](PowerPoint.Slides.FindBySlideID.md)** method with the slide's ID number can be a more reliable way to return a specific **Slide** object from a **Slides** collection than using the **Item** method with the slide's index number.


## Example

This example demonstrates how to retrieve the unique ID number for a  **Slide** object and then use this number to return that **Slide** object from the **Slides** collection.


```vb
Set gslides = ActivePresentation.Slides

'Get slide ID
graphSlideID = gslides.Add(2, ppLayoutChart).SlideID

gslides.FindBySlideID(graphSlideID) _
    .SlideShowTransition.EntryEffect = _
    ppEffectCoverLeft      'Use ID to return specific slide
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]