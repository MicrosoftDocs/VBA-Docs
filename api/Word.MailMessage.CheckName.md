---
title: MailMessage.CheckName method (Word)
keywords: vbawd10.chm163184974
f1_keywords:
- vbawd10.chm163184974
ms.prod: word
api_name:
- Word.MailMessage.CheckName
ms.assetid: 2888dfb7-5773-cbf8-8865-c90875411476
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.CheckName method (Word)

Validates the email addresses that appear in the  **To**,  **Cc**, and  **Bcc** lines in the active email message.


## Syntax

_expression_. `CheckName`

_expression_ Required. A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Remarks

This method is available only if you are using Word as your email editor. If the names cannot be validated, the  **Check Names** dialog box is displayed.


## Example

This example validates the email addresses that appear in the active email message.


```vb
Application.MailMessage.CheckName
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]