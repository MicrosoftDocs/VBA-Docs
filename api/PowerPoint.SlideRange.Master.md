---
title: SlideRange.Master property (PowerPoint)
keywords: vbapp10.chm532023
f1_keywords:
- vbapp10.chm532023
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Master
ms.assetid: 321cb5f9-2ac8-f31c-2c79-0cfdc4e0a73b
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Master property (PowerPoint)

Returns a  **[Master](PowerPoint.Master.md)** object that represents the slide master. Read-only.


## Syntax

_expression_. `Master`

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Master


## Example

This example sets the background fill for the slide master for slide one in the active presentation.


```vb
ActivePresentation.Slides(1).Master.Background.Fill _
    .PresetGradient msoGradientDiagonalUp, 1, msoGradientDaybreak
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]