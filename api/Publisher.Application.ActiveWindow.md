---
title: Application.ActiveWindow property (Publisher)
keywords: vbapb10.chm131074
f1_keywords:
- vbapb10.chm131074
ms.prod: publisher
api_name:
- Publisher.Application.ActiveWindow
ms.assetid: 125e2bb4-f922-ceef-9e3e-5dbe3aaff2a4
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.ActiveWindow property (Publisher)

Returns a **[Window](Publisher.Window.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.


## Syntax

_expression_.**ActiveWindow**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Example

This example displays the active window's caption.

```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]