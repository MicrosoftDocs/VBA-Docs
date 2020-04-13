---
title: ContactItem.Email3Address property (Outlook)
keywords: vbaol11.chm999
f1_keywords:
- vbaol11.chm999
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email3Address
ms.assetid: b0f29077-a06c-a2cf-e873-b9d560d91498
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Email3Address property (Outlook)

Returns or sets a **String** representing the email address of the third email entry for the contact. Read/write.


## Syntax

_expression_. `Email3Address`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the email address for the third email entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email3Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]