---
title: Slide.ColorScheme property (PowerPoint)
keywords: vbapp10.chm531006
f1_keywords:
- vbapp10.chm531006
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.ColorScheme
ms.assetid: 3d40d93f-4e7d-e95f-8340-d138da2a1b55
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.ColorScheme property (PowerPoint)

Returns or sets the  **[ColorScheme](PowerPoint.ColorScheme.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.


## Syntax

_expression_. `ColorScheme`

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

ColorScheme


## Example

This example sets the title color to green for slides one and three in the active presentation.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]