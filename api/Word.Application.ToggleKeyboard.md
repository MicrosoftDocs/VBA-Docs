---
title: Application.ToggleKeyboard method (Word)
keywords: vbawd10.chm158335378
f1_keywords:
- vbawd10.chm158335378
ms.prod: word
api_name:
- Word.Application.ToggleKeyboard
ms.assetid: a7af90f6-28e5-6655-ae5b-c01ed64da52f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ToggleKeyboard method (Word)

Switches the keyboard language setting between right-to-left and left-to-right languages.


## Syntax

_expression_. `ToggleKeyboard`

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example asks the user whether to switch the keyboard language setting between right-to-left and left-to-right languages.


```vb
x = MsgBox("Switch the keyboard language setting?", vbYesNo) 
If x = vbYes Then Application.ToggleKeyboard
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]