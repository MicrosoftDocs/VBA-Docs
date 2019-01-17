---
title: NamedSlideShows.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShows.Count
ms.assetid: e4a48f6c-32f8-fdc5-101d-3ddec1f79f59
ms.date: 06/08/2017
localization_priority: Normal
---


# NamedSlideShows.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. `Count`

 _expression_ A variable that represents a [NamedSlideShows](./PowerPoint.NamedSlideShows.md) object.


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


[NamedSlideShows Object](PowerPoint.NamedSlideShows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]