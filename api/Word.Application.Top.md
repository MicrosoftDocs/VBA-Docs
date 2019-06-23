---
title: Application.Top property (Word)
keywords: vbawd10.chm158335064
f1_keywords:
- vbawd10.chm158335064
ms.prod: word
api_name:
- Word.Application.Top
ms.assetid: bbce9fe2-8390-f73d-8fca-bd047df468be
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Top property (Word)

Returns or sets the vertical position of the active document. Read/write  **Long**.


## Syntax

_expression_.**Top**

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example positions the Word application window 100 points from the top of the screen.


```vb
Application.WindowState = wdWindowStateNormal 
Application.Top = 100
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]