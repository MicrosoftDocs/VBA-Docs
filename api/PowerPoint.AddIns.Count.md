---
title: AddIns.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.AddIns.Count
ms.assetid: 5ccbf78a-3585-8de5-78c9-b27f32d8f5c9
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an [AddIns](PowerPoint.AddIns.md) object.


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


[AddIns Object](PowerPoint.AddIns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]