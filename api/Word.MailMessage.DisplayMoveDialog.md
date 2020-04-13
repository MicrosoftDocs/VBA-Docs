---
title: MailMessage.DisplayMoveDialog method (Word)
keywords: vbawd10.chm163184976
f1_keywords:
- vbawd10.chm163184976
ms.prod: word
api_name:
- Word.MailMessage.DisplayMoveDialog
ms.assetid: e913a4f3-e970-ae2f-84b1-c239cc57a15f
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.DisplayMoveDialog method (Word)

Displays the **Move** dialog box, in which the user can specify a new location for the active email message in an available message store.


## Syntax

_expression_. `DisplayMoveDialog`

_expression_ Required. A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Remarks

This method is available only if you are using Word as your email editor.


## Example

This example displays the **Move** dialog box for the active email message.


```vb
Application.MailMessage.DisplayMoveDialog
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]