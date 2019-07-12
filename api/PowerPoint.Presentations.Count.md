---
title: Presentations.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.Count
ms.assetid: e9f4d85f-4ba3-6c07-353d-79bbf39f91da
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentations.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Presentations](PowerPoint.Presentations.md) object.


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


[Presentations Object](PowerPoint.Presentations.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]