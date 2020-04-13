---
title: MailMessage.DisplaySelectNamesDialog method (Word)
keywords: vbawd10.chm163184978
f1_keywords:
- vbawd10.chm163184978
ms.prod: word
api_name:
- Word.MailMessage.DisplaySelectNamesDialog
ms.assetid: 54b3d2fd-42db-a4da-4247-cc0b0eca5f65
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage.DisplaySelectNamesDialog method (Word)

Displays the **Select Names** dialog box, in which the user can add addresses to the **To**,  **Cc**, and  **Bcc** lines in the active, unsent email message.


## Syntax

_expression_. `DisplaySelectNamesDialog`

_expression_ Required. A variable that represents a '[MailMessage](Word.MailMessage.md)' object.


## Remarks

This method is available only if you are using Word as your email editor.


## Example

This example displays the **Select Names** dialog box for the active email message.


```vb
Application.MailMessage.DisplaySelectNamesDialog
```


## See also


[MailMessage Object](Word.MailMessage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]