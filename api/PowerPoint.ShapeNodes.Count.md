---
title: ShapeNodes.Count property (PowerPoint)
keywords: vbapp10.chm560002
f1_keywords:
- vbapp10.chm560002
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes.Count
ms.assetid: 63f8a4da-1b0a-b72c-06ed-27477fb74809
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeNodes.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[ShapeNodes](PowerPoint.ShapeNodes.md)** object.


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


[ShapeNodes Object](PowerPoint.ShapeNodes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]