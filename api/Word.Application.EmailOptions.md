---
title: Application.EmailOptions property (Word)
keywords: vbawd10.chm158335365
f1_keywords:
- vbawd10.chm158335365
ms.prod: word
api_name:
- Word.Application.EmailOptions
ms.assetid: 28547346-6119-b763-339e-b04af1c8268f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EmailOptions property (Word)

Returns an  **[EmailOptions](Word.EmailOptions.md)** object that represents the global preferences for email authoring. Read-only.


## Syntax

_expression_. `EmailOptions`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example sets Microsoft Word to mark comments in email messages.


```vb
Application.EmailOptions.MarkComments = True
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]