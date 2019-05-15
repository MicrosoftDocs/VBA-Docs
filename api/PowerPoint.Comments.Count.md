---
title: Comments.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Comments.Count
ms.assetid: b03db1bc-f969-8a27-bfd2-4327e699c08a
ms.date: 06/08/2017
localization_priority: Normal
---


# Comments.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[Comments](PowerPoint.Comments.md)** object.


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


[Comments Object](PowerPoint.Comments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]