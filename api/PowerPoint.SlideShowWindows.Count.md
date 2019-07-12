---
title: SlideShowWindows.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindows.Count
ms.assetid: 19f91cd6-c12d-92b1-21e9-a3a0916bf4df
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowWindows.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [SlideShowWindows](PowerPoint.SlideShowWindows.md) object.


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


[SlideShowWindows Object](PowerPoint.SlideShowWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]