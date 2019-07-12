---
title: Panes.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Panes.Count
ms.assetid: 450fb25b-46b5-00e5-4e26-f08974ca14e0
ms.date: 06/08/2017
localization_priority: Normal
---


# Panes.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Panes](PowerPoint.Panes.md) object.


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


[Panes Object](PowerPoint.Panes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]