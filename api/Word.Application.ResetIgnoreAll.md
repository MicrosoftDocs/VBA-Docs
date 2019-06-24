---
title: Application.ResetIgnoreAll method (Word)
keywords: vbawd10.chm158335302
f1_keywords:
- vbawd10.chm158335302
ms.prod: word
api_name:
- Word.Application.ResetIgnoreAll
ms.assetid: 8a6dcb30-23bb-70bb-e257-e519bc63a289
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ResetIgnoreAll method (Word)

Clears the list of words that were previously ignored during a spelling check.


## Syntax

_expression_. `ResetIgnoreAll`

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

 After you run this method, previously ignored words are checked along with all the other words. In order for the **ResetIgnoreAll** method to work, you must first set the **SpellingChecked** property to **False**.


## Example

This example clears the list of words that were ignored during a previous spelling check and then begins a new spelling check on the active document.


```vb
Application.ResetIgnoreAll 
ActiveDocument.SpellingChecked = False 
ActiveDocument.CheckSpelling
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]