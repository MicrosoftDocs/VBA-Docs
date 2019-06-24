---
title: Application.NumLock property (Word)
keywords: vbawd10.chm158335025
f1_keywords:
- vbawd10.chm158335025
ms.prod: word
api_name:
- Word.Application.NumLock
ms.assetid: 0c20c000-2df9-1483-91be-cacf1abe0ff0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NumLock property (Word)

Returns the state of the NUM LOCK key.  **True** if the keys on the numeric keypad insert numbers, **False** if the keys move the insertion point. Read-only **Boolean**.


## Syntax

_expression_. `NumLock`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

This example returns the current state of the NUM LOCK key.


```vb
theState = Application.NumLock
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]