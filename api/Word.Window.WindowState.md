---
title: Window.WindowState property (Word)
keywords: vbawd10.chm157417483
f1_keywords:
- vbawd10.chm157417483
ms.prod: word
api_name:
- Word.Window.WindowState
ms.assetid: 0be17839-28d5-6ba7-5f66-02504a4aa604
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.WindowState property (Word)

Returns or sets the state of the specified document window or task window. Read/write  **[WdWindowState](Word.WdWindowState.md)**.


## Syntax

_expression_.**WindowState**

_expression_ Required. A variable that represents a **[Window](Word.Window.md)** object.


## Remarks

The **wdWindowStateNormal** constant indicates a window that's not maximized or minimized. The state of an inactive window cannot be set. Use the **[Activate](Word.Window.Activate.md)** method to activate a window prior to setting the window state.


## Example

This example maximizes the active window if it is not maximized or minimized.


```vb
If ActiveDocument.ActiveWindow _ 
 .WindowState = wdWindowStateNormal Then _ 
 ActiveDocument.ActiveWindow _ 
 .WindowState = wdWindowStateMaximize
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]