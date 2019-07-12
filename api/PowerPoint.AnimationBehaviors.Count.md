---
title: AnimationBehaviors.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors.Count
ms.assetid: cffdda44-6b03-b25f-b21a-8e0e350d5d87
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehaviors.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an [AnimationBehaviors](PowerPoint.AnimationBehaviors.md) object.


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


[AnimationBehaviors Object](PowerPoint.AnimationBehaviors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]