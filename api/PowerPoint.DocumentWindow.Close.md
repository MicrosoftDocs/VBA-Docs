---
title: DocumentWindow.Close method (PowerPoint)
keywords: vbapp10.chm511020
f1_keywords:
- vbapp10.chm511020
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.Close
ms.assetid: c7ba0097-5fa3-b0d0-234b-3cfe3e493522
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.Close method (PowerPoint)

Closes the specified document window.


## Syntax

_expression_.**Close**

_expression_ A variable that represents a [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Remarks

When you use this method, PowerPoint will close an open presentation without prompting users to save their work. To prevent the loss of work, use the  **Save** method or the **SaveAs** method before you use the **Close** method.


## Example

This example closes all windows except the active window.


```vb
With Application.Windows

    For i = 2 To .Count

        .Item(i).Close

    Next

End With
```


## See also



[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]