---
title: Presentation.Slides property (PowerPoint)
keywords: vbapp10.chm583011
f1_keywords:
- vbapp10.chm583011
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Slides
ms.assetid: bf481c73-3508-a074-eb2c-a5df62e55a5c
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Slides property (PowerPoint)

Returns a  **[Slides](PowerPoint.Slides.md)** collection that represents all slides in the specified presentation. Read-only.


## Syntax

_expression_. `Slides`

_expression_ A variable that represents a [PlaySettings](PowerPoint.PlaySettings.md) object.


## Return value

Slides


## Example

This example adds a slide to the active presentation.


```vb
Application.ActivePresentation.Slides.Add 1, ppLayoutTitle
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]