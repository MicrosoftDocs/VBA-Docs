---
title: RulerLevels.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.RulerLevels.Count
ms.assetid: 5278b041-dabb-7b14-32ef-528b238d3326
ms.date: 06/08/2017
localization_priority: Normal
---


# RulerLevels.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [RulerLevels](PowerPoint.RulerLevels.md) object.


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


[RulerLevels Object](PowerPoint.RulerLevels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]