---
title: SlideRange.Count property (PowerPoint)
keywords: vbapp10.chm532029
f1_keywords:
- vbapp10.chm532029
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Count
ms.assetid: ff280d05-41f1-fbdc-16c3-ae30a1102340
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Return value

Long


## Example

This example closes all windows except the active window.


```vb
With Application.Windows

    For i = 2 To .Count

        .Item(2).Close

    Next

End With
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]