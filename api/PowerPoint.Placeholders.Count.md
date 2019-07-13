---
title: Placeholders.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Placeholders.Count
ms.assetid: 8f20feee-b574-a5f1-1499-655495056178
ms.date: 06/08/2017
localization_priority: Normal
---


# Placeholders.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [Placeholders](PowerPoint.Placeholders.md) object.


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


[Placeholders Object](PowerPoint.Placeholders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]