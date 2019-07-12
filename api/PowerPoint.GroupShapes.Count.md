---
title: GroupShapes.Count property (PowerPoint)
keywords: vbapp10.chm549002
f1_keywords:
- vbapp10.chm549002
ms.prod: powerpoint
api_name:
- PowerPoint.GroupShapes.Count
ms.assetid: 1535f42d-43de-a9e2-0972-a1bb34b5578f
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupShapes.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [GroupShapes](PowerPoint.GroupShapes.md) object.


## Return value

[INT]


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


[GroupShapes Object](PowerPoint.GroupShapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]