---
title: SlideRange.Hyperlinks property (PowerPoint)
keywords: vbapp10.chm532024
f1_keywords:
- vbapp10.chm532024
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Hyperlinks
ms.assetid: bfa4da43-4c56-e010-0848-2cb55fb68154
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Hyperlinks property (PowerPoint)

Returns a  **[Hyperlinks](PowerPoint.Hyperlinks.md)** collection that represents all the hyperlinks on the specified slide. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Hyperlinks


## Example

This example allows the user to update an outdated Internet address for all hyperlinks in the active presentation.


```vb
oldAddr = InputBox("Old Internet address")

newAddr = InputBox("New Internet address")

For Each s In ActivePresentation.Slides

    For Each h In s.Hyperlinks

        If LCase(h.Address) = oldAddr Then h.Address = newAddr

    Next

Next
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]