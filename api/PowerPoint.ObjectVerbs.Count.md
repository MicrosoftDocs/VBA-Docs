---
title: ObjectVerbs.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ObjectVerbs.Count
ms.assetid: 8aabdb50-1e4a-655b-5336-5ae7be5a65b1
ms.date: 06/08/2017
localization_priority: Normal
---


# ObjectVerbs.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an [ObjectVerbs](PowerPoint.ObjectVerbs.md) object.


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


[ObjectVerbs Object](PowerPoint.ObjectVerbs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]