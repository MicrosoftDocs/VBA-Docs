---
title: Sequences.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Sequences.Count
ms.assetid: 3292024f-d87d-8031-29ab-11631361cd99
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequences.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Sequences](PowerPoint.Sequences.md) object.


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


[Sequences Object](PowerPoint.Sequences.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]