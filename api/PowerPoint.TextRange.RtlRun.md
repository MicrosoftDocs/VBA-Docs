---
title: TextRange.RtlRun method (PowerPoint)
keywords: vbapp10.chm569038
f1_keywords:
- vbapp10.chm569038
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.RtlRun
ms.assetid: eb474c9b-d789-f741-9ba9-0514f0a5b0be
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.RtlRun method (PowerPoint)

Sets the direction of text in a text range to read from right to left.


## Syntax

_expression_. `RtlRun`

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Remarks

This method makes it possible to use text from both left-to-right and right-to-left languages in the same presentation.


## Example

The following example finds all of the shapes on slide one that contain text and changes the text to read from right to left.


```vb
ActiveWindow.ViewType = ppViewSlide

For Each sh In ActivePresentation.Slides(1).Shapes

    If sh.HasTextFrame Then

         sh.TextFrame.TextRange.RtlRun

    End If

Next
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]