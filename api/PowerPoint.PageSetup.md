---
title: PageSetup object (PowerPoint)
keywords: vbapp10.chm527000
f1_keywords:
- vbapp10.chm527000
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup
ms.assetid: aed5649c-59d7-08d2-0a01-3385e5a9b5ff
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup object (PowerPoint)

Contains information about the page setup for slides, notes pages, handouts, and outlines in a presentation.


## Example

Use the [PageSetup](PowerPoint.Presentation.PageSetup.md)property to return the  **PageSetup** object. The following example sets all slides in the active presentation to be 11 inches wide and 8.5 inches high and sets the slide numbering for the presentation to start at 17.


```vb
With ActivePresentation.PageSetup

    .SlideWidth = 11 * 72

    .SlideHeight = 8.5 * 72

    .FirstSlideNumber = 17

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]