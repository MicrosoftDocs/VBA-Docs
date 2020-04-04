---
title: Master.Hyperlinks property (PowerPoint)
keywords: vbapp10.chm533012
f1_keywords:
- vbapp10.chm533012
ms.prod: powerpoint
api_name:
- PowerPoint.Master.Hyperlinks
ms.assetid: 5d9af48b-49e2-4253-a431-4341a697437b
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.Hyperlinks property (PowerPoint)

Returns a **[Hyperlinks](PowerPoint.Hyperlinks.md)** collection that represents all the hyperlinks on the specified slide. Read-only.


## Syntax

_expression_.**Hyperlinks**

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


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


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]