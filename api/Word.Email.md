---
title: Email object (Word)
keywords: vbawd10.chm2525
f1_keywords:
- vbawd10.chm2525
ms.prod: word
api_name:
- Word.Email
ms.assetid: ee23a74e-556b-04d8-f0b9-fb95f7aa8cfc
ms.date: 06/08/2017
localization_priority: Normal
---


# Email object (Word)

Represents an email message.


## Remarks

Use the **[Email](Word.Document.Email.md)** property to return the **Email** object. The **Email** object and its properties are valid only if the active document is an unsent forward, reply, or new email message.

This example displays the name of the style associated with the current email author.




```vb
MsgBox ActiveDocument.Email _ 
 .CurrentEmailAuthor.Style.NameLocal
```

The author style name is the same as the value returned by the **[UserName](Word.Application.UserName.md)** property.


> [!NOTE] 
>  There is no Emails collection; each **Document** object contains only one **Email** object.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]