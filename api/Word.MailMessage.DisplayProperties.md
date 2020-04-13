---
title: MailMessage.DisplayProperties method (Word)
keywords: vbawd10.chm163184977
f1_keywords:
- vbawd10.chm163184977
ms.prod: word
api_name:
- Word.MailMessage.DisplayProperties
ms.assetid: fa660e11-5329-5167-ddc3-0d90ee820251
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.DisplayProperties method (Word)

Displays the **Properties** dialog box for the active email message.


## Syntax

_expression_. `DisplayProperties`

_expression_ Required. A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Remarks

This method is available only if you are using Word as your email editor.


## Example

This example displays the **Properties** dialog box for the active email message.


```vb
Application.MailMessage.DisplayProperties
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]