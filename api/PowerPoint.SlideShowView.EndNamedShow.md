---
title: SlideShowView.EndNamedShow method (PowerPoint)
keywords: vbapp10.chm513023
f1_keywords:
- vbapp10.chm513023
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.EndNamedShow
ms.assetid: 1b829558-a729-8aa1-c260-8b7410501153
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.EndNamedShow method (PowerPoint)

Switches from running a custom, or named, slide show to running the entire presentation of which the custom show is a subset. When the slide show advances from the current slide, the next slide displayed will be the next one in the entire presentation, not the next one in the custom slide show.


## Syntax

_expression_. `EndNamedShow`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Example

If a custom slide show is currently running in slide show window one, this example redefines the slide show to include all the slides in the presentation from which the slides in the custom show were selected.


```vb
SlideShowWindows(1).View.EndNamedShow
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]