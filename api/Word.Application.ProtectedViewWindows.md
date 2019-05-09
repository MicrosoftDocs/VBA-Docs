---
title: Application.ProtectedViewWindows property (Word)
keywords: vbawd10.chm158335466
f1_keywords:
- vbawd10.chm158335466
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindows
ms.assetid: eb1c8cae-c0da-0a84-316e-808302869b26
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindows property (Word)

Returns a [ProtectedViewWindows](Word.ProtectedViewWindows.md) collection that represents all Protected View windows. Read-only.


## Syntax

_expression_. `ProtectedViewWindows`

 _expression_ An expression that returns an [Application](./Word.Application.md) object.


## Example

The following code example displays the number of Protected View windows that are open.


```vb
MsgBox "There are " & ProtectedViewWindows.Count & _ 
 " Protected View windows open."
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]