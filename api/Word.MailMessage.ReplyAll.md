---
title: MailMessage.ReplyAll method (Word)
keywords: vbawd10.chm163184983
f1_keywords:
- vbawd10.chm163184983
ms.prod: word
api_name:
- Word.MailMessage.ReplyAll
ms.assetid: cc7aa537-573f-f2b2-14a1-3443ed622f56
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.ReplyAll method (Word)

Opens a new email message — with the sender's and all other recipients' addresses on the **To** and **Cc** lines, as appropriate — for replying to the active message.


## Syntax

_expression_. `ReplyAll`

_expression_ Required. A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Example

This example opens a new email message for replying to the active message.


```vb
Application.MailMessage.ReplyAll
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]