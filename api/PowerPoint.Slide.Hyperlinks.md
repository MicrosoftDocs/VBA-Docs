---
title: Slide.Hyperlinks property (PowerPoint)
keywords: vbapp10.chm531024
f1_keywords:
- vbapp10.chm531024
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Hyperlinks
ms.assetid: 0e1d7545-815f-3be9-38b8-355f9e6e9962
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Hyperlinks property (PowerPoint)

Returns a **[Hyperlinks](PowerPoint.Hyperlinks.md)** collection that represents all the hyperlinks on the specified slide. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


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


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]