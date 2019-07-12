---
title: NamedSlideShow.Count property (PowerPoint)
keywords: vbapp10.chm516006
f1_keywords:
- vbapp10.chm516006
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShow.Count
ms.assetid: 09aeed71-dfc6-2ee6-1430-c5e7f0ed2bc1
ms.date: 06/08/2017
localization_priority: Normal
---


# NamedSlideShow.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [NamedSlideShow](PowerPoint.NamedSlideShow.md) object.


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


[NamedSlideShow Object](PowerPoint.NamedSlideShow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]