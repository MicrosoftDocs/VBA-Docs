---
title: HeadersFooters.SlideNumber property (PowerPoint)
keywords: vbapp10.chm542004
f1_keywords:
- vbapp10.chm542004
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters.SlideNumber
ms.assetid: c846069f-dd3f-c5ac-f9ac-b5a7ed499bdc
ms.date: 06/08/2017
localization_priority: Normal
---


# HeadersFooters.SlideNumber property (PowerPoint)

Returns a  **[HeaderFooter](PowerPoint.HeaderFooter.md)** object that represents the slide number in the lower-right corner of a slide, or the page number in the lower-right corner of a notes page or a page of a printed handout or outline. Read-only.


## Syntax

_expression_. `SlideNumber`

_expression_ A variable that represents a [HeadersFooters](PowerPoint.HeadersFooters.md) object.


## Return value

HeaderFooter


## Example

This example hides the slide number on slide two in the active presentation if the number is currently visible, or it displays the slide number if it is currently hidden.


```vb
With Application.ActivePresentation.Slides(2) _
        .HeadersFooters.SlideNumber
    If .Visible Then
        .Visible = False
    Else
        .Visible = True
    End If
End With
```


## See also


[HeadersFooters Object](PowerPoint.HeadersFooters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]