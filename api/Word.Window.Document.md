---
title: Window.Document property (Word)
keywords: vbawd10.chm157417474
f1_keywords:
- vbawd10.chm157417474
ms.prod: word
api_name:
- Word.Window.Document
ms.assetid: a1eda09e-9c5b-548a-23d0-27cbda9e0dcd
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Document property (Word)

Returns a  **[Document](Word.Document.md)** object associated with the specified pane, window, or selection. Read-only.


## Syntax

_expression_.**Document**

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example sets myDoc to the document associated with the active window. The focus is changed to the next window, and the window is split. The **Activate** method is used to switch back to the original document.


```vb
Set myDoc = Application.ActiveWindow.Document 
If Windows.Count >= 2 Then 
 Application.ActiveWindow.Next.Activate 
 Application.ActiveWindow.Split = True 
 myDoc.Activate 
End If
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]