---
title: Sequence.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.Count
ms.assetid: b3f02a35-309d-768c-dc76-bd0ef84261cc
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Sequence](PowerPoint.Sequence.md) object.


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


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]