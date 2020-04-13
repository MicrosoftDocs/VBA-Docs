---
title: ProtectedViewWindow.WindowState property (Word)
keywords: vbawd10.chm231735302
f1_keywords:
- vbawd10.chm231735302
ms.prod: word
api_name:
- Word.ProtectedViewWindow.WindowState
ms.assetid: 858036ac-b0f2-f312-81f2-d02ace00b2fb
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.WindowState property (Word)

Returns or sets the state of the specified Protected View window. Read/write [WdWindowState](Word.WdWindowState.md).


## Syntax

_expression_.**WindowState**

 _expression_ An expression that returns a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Remarks

The **wdWindowStateNormal** constant indicates a window that is not maximized or minimized. The state of an inactive window cannot be set. Use the [Activate](Word.ProtectedViewWindow.Activate.md) method to activate a window prior to setting the window state.


## Example

The following code example minimizes the active Protected View window.


```vb
ActiveProtectedViewWindow.WindowState = wdWindowStateMinimize
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]