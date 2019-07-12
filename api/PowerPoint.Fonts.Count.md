---
title: Fonts.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Fonts.Count
ms.assetid: 94f6cfda-23f5-0a89-388f-6cb3b544fdb6
ms.date: 06/08/2017
localization_priority: Normal
---


# Fonts.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Fonts](PowerPoint.Fonts.md) object.


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


[Fonts Object](PowerPoint.Fonts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]