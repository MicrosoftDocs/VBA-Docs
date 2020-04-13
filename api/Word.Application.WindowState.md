---
title: Application.WindowState property (Word)
keywords: vbawd10.chm158335067
f1_keywords:
- vbawd10.chm158335067
ms.prod: word
api_name:
- Word.Application.WindowState
ms.assetid: ae457f42-9c12-d0f4-e74e-d01610b9b4af
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowState property (Word)

Returns or sets the state of the specified document window or task window. Read/write  **WdWindowState**.


## Syntax

_expression_.**WindowState**

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

The **wdWindowStateNormal** constant indicates a window that's not maximized or minimized. The state of an inactive window cannot be set. Use the **Activate** method to activate a window prior to setting the window state.


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]