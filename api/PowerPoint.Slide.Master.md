---
title: Slide.Master property (PowerPoint)
keywords: vbapp10.chm531023
f1_keywords:
- vbapp10.chm531023
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Master
ms.assetid: cec5385d-f6af-dd8d-7989-251a70c4937e
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Master property (PowerPoint)

Returns a  **[Master](PowerPoint.Master.md)** object that represents the slide master. Read-only.


## Syntax

_expression_. `Master`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

Master


## Example

This example sets the background fill for the slide master for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Master.Background.Fill _
    .PresetGradient msoGradientDiagonalUp, 1, msoGradientDaybreak
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]