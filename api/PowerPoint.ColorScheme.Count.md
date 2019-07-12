---
title: ColorScheme.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ColorScheme.Count
ms.assetid: 372e48be-db37-82a1-8bca-1ac71b6ae165
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorScheme.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [ColorScheme](PowerPoint.ColorScheme.md) object.


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


[ColorScheme Object](PowerPoint.ColorScheme.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]