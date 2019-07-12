---
title: Selection.SlideRange property (PowerPoint)
keywords: vbapp10.chm508008
f1_keywords:
- vbapp10.chm508008
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.SlideRange
ms.assetid: 2d853875-b0c2-ab8e-38b6-4e1397d4e669
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SlideRange property (PowerPoint)

Returns a  **[SlideRange](PowerPoint.SlideRange.md)** object that represents a range of selected slides. Read-only.


## Syntax

_expression_. `SlideRange`

_expression_ A variable that represents a [Selection](PowerPoint.Selection.md) object.


## Return value

SlideRange


## Remarks

A slide range can be constructed in slide view, slide sorter view, normal view, notes page view, or outline view. In slide view,  **SlideRange** returns one slide — the current, displayed slide.


## Example

This example sets the background scheme color for all the selected slides in window one.


```vb
Windows(1).Selection.SlideRange.ColorScheme _
    .Colors(ppBackground).RGB = RGB(0, 0, 255)
```


## See also


[Selection Object](PowerPoint.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]