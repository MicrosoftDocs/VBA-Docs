---
title: Master.ColorScheme property (PowerPoint)
keywords: vbapp10.chm533005
f1_keywords:
- vbapp10.chm533005
ms.prod: powerpoint
api_name:
- PowerPoint.Master.ColorScheme
ms.assetid: f481aa76-e96f-686a-edbb-b2bef8be0e8c
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.ColorScheme property (PowerPoint)

Returns or sets the  **[ColorScheme](PowerPoint.ColorScheme.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.


## Syntax

_expression_. `ColorScheme`

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


## Return value

ColorScheme


## Example

This example sets the title color to green for slides one and three in the active presentation.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```


## See also


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]