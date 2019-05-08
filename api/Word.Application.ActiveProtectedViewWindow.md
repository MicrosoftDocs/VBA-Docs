---
title: Application.ActiveProtectedViewWindow property (Word)
keywords: vbawd10.chm158335467
f1_keywords:
- vbawd10.chm158335467
ms.prod: word
api_name:
- Word.Application.ActiveProtectedViewWindow
ms.assetid: 2ba10f3d-3f43-5628-a5fc-3c65b290ef72
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveProtectedViewWindow property (Word)

Returns a [ProtectedViewWindow](Word.ProtectedViewWindow.md) object that represents the active Protected View window. Read-only.


## Syntax

_expression_. `ActiveProtectedViewWindow`

 _expression_ An expression that returns a [Application](./Word.Application.md) object.


## Example

The following code example displays the caption text for the active Protected View window.


```vb
MsgBox ActiveProtectedViewWindow.Caption 

```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]