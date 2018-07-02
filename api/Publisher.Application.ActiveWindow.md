---
title: Application.ActiveWindow Property (Publisher)
keywords: vbapb10.chm131074
f1_keywords:
- vbapb10.chm131074
ms.prod: publisher
api_name:
- Publisher.Application.ActiveWindow
ms.assetid: 125e2bb4-f922-ceef-9e3e-5dbe3aaff2a4
ms.date: 06/08/2017
---


# Application.ActiveWindow Property (Publisher)

Returns a  **[Window](Publisher.Window.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.


## Syntax

 _expression_. **ActiveWindow**

 _expression_ A variable that represents an  **Application** object.


## Example

This example displays the active window's caption.


```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```


## See also


 [Application Object](Publisher.Application.md)

