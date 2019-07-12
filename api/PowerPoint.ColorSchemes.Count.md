---
title: ColorSchemes.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ColorSchemes.Count
ms.assetid: bae2f5a0-094a-cffb-af36-9ce8c042fde8
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorSchemes.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [ColorSchemes](PowerPoint.ColorSchemes.md) object.


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


[ColorSchemes Object](PowerPoint.ColorSchemes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]