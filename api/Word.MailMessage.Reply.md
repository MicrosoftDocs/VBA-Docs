---
title: MailMessage.Reply method (Word)
keywords: vbawd10.chm163184982
f1_keywords:
- vbawd10.chm163184982
ms.prod: word
api_name:
- Word.MailMessage.Reply
ms.assetid: a05e3352-84bb-8774-c841-d2b6093dcf9b
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.Reply method (Word)

Opens a new email message — with the sender's address on the **To** line — for replying to the active message.


## Syntax

_expression_. `Reply`

_expression_ A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Example

This example opens a new email message for replying to the active message.


```vb
Application.MailMessage.Reply
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]