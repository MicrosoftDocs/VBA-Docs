---
title: Application.Version property (Word)
keywords: vbawd10.chm158335000
f1_keywords:
- vbawd10.chm158335000
ms.prod: word
api_name:
- Word.Application.Version
ms.assetid: 7bdd0acc-1ed0-677c-f973-99a9199e030b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Version property (Word)

Returns the Microsoft Word version number. Read-only  **String**.


## Syntax

_expression_.**Version**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example displays the Word version number in a message box.


```vb
Msgbox "The version of Word is " & Application.Version
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]