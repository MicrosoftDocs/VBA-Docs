---
title: DocumentWindows.Count property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows.Count
ms.assetid: d659a980-cc23-c805-6084-4c724c0bc6cd
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindows.Count property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a [DocumentWindows](PowerPoint.DocumentWindows.md) object.


## Return value

Long


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.DocumentWindows.Count**
    

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



[DocumentWindows Object](PowerPoint.DocumentWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]