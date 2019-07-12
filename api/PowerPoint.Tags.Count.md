---
title: Tags.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Tags.Count
ms.assetid: 4a6ae9cb-65f8-c273-e50c-e75d6a785767
ms.date: 06/08/2017
localization_priority: Normal
---


# Tags.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Tags](PowerPoint.Tags.md) object.


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


[Tags Object](PowerPoint.Tags.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]