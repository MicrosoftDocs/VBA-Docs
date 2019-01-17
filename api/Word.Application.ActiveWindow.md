---
title: Application.ActiveWindow property (Word)
keywords: vbawd10.chm158334980
f1_keywords:
- vbawd10.chm158334980
ms.prod: word
api_name:
- Word.Application.ActiveWindow
ms.assetid: 0a376a89-7bba-d5fd-8073-7cb2e79a87a8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveWindow property (Word)

Returns a  **[Window](Word.Window.md)** object that represents the active window (the window with the focus). If there are no windows open, an error occurs. Read-only.


## Syntax

 _expression_. `ActiveWindow`

 _expression_ A variable that represents an '[Application](Word.Application.md)' object.


## Example

This example displays the caption text for the active window.


```vb
Sub WindowCaption() 
 MsgBox ActiveWindow.Caption 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]