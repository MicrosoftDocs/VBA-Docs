---
title: Document.ActiveWindow property (Publisher)
keywords: vbapb10.chm196611
f1_keywords:
- vbapb10.chm196611
ms.prod: publisher
api_name:
- Publisher.Document.ActiveWindow
ms.assetid: 0d00a8fa-aef2-43df-3c54-0cca804b7eee
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ActiveWindow property (Publisher)

Returns a **[Window](Publisher.Window.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.


## Syntax

_expression_.**ActiveWindow**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Example

This example displays the active window's caption.

```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]