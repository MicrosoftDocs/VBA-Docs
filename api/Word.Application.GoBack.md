---
title: Application.GoBack method (Word)
keywords: vbawd10.chm158335304
f1_keywords:
- vbawd10.chm158335304
ms.prod: word
api_name:
- Word.Application.GoBack
ms.assetid: d1113bc7-4ad3-f4da-0442-c11f5e22b2a8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GoBack method (Word)

Moves the insertion point among the last three locations where editing occurred in the active document (the same as pressing SHIFT+F5).


## Syntax

_expression_. `GoBack`

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example opens the most recently used file and then moves the insertion point to the location where editing last occurred.


```vb
RecentFiles(1).Open 
Application.GoBack
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]