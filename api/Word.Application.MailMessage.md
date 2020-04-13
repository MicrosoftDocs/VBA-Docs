---
title: Application.MailMessage property (Word)
keywords: vbawd10.chm158335324
f1_keywords:
- vbawd10.chm158335324
ms.prod: word
api_name:
- Word.Application.MailMessage
ms.assetid: 82bca039-0b6b-4489-27bf-18746dc639d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMessage property (Word)

Returns a  **[MailMessage](Word.MailMessage.md)** object that represents the active email message.


## Syntax

_expression_. `MailMessage`

 _expression_ An expression that returns an **[Application](Word.Application.md)** object. 


## Example

This example displays the **Select Names** dialog box for the active email message.


```vb
Application.MailMessage.DisplaySelectNamesDialog
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]