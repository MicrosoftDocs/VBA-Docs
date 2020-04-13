---
title: EmailAuthor object (Word)
keywords: vbawd10.chm2519
f1_keywords:
- vbawd10.chm2519
ms.prod: word
api_name:
- Word.EmailAuthor
ms.assetid: 2749e018-42e9-7a1a-f18b-8605b38ff0ae
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailAuthor object (Word)

Represents the author of an email message.


## Remarks

Use the **[CurrentEmailAuthor](Word.Email.CurrentEmailAuthor.md)** property to return the **EmailAuthor** object. The **EmailAuthor** object and its properties are valid only if the active document is an unsent forward, reply, or new email message.

This example returns the style associated with the current author for unsent replies, forwards, or new email messages, and displays the name of the font associated with this style.




```vb
Set MyEmailStyle = _ 
 ActiveDocument.Email.CurrentEmailAuthor.Style 
Msgbox MyEmailStyle.Font.Name
```


> [!NOTE] 
> There is no EmailAuthors collection; each  **Email** object contains only one **EmailAuthor** object.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]