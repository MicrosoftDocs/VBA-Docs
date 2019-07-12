---
title: Columns.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Columns.Count
ms.assetid: d23ac7d2-080f-9981-b502-16ba11d811e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Columns](PowerPoint.Columns.md) object.


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


[Columns Object](PowerPoint.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]