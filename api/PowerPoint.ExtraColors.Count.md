---
title: ExtraColors.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ExtraColors.Count
ms.assetid: a3e5efcd-498f-3c29-2577-25cdc7206a7b
ms.date: 06/08/2017
localization_priority: Normal
---


# ExtraColors.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an [ExtraColors](PowerPoint.ExtraColors.md) object.


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


[ExtraColors Object](PowerPoint.ExtraColors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]