---
title: Slides.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.Count
ms.assetid: b01d04ed-b28f-608e-b77f-2ef94e1a2d2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Slides.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Slides](PowerPoint.Slides.md) object.


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


[Slides Object](PowerPoint.Slides.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]