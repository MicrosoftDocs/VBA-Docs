---
title: MailMessage object (Word)
keywords: vbawd10.chm2490
f1_keywords:
- vbawd10.chm2490
ms.prod: word
api_name:
- Word.MailMessage
ms.assetid: d0109969-27f7-0180-c56d-5b49a3f0171b
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMessage object (Word)

Represents the active email message if you are using Microsoft Word as your email editor.


## Remarks

Use the **MailMessage** property to return the **MailMessage** object. The following example validates the email addresses that appear in the active email message.


```vb
Application.MailMessage.CheckName
```

The methods of the **MailMessage** object require that you are using Word as your email editor and that an email message is active. If either of these conditions is not true, an error occurs.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]