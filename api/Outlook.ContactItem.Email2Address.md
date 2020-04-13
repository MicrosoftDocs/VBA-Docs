---
title: ContactItem.Email2Address property (Outlook)
keywords: vbaol11.chm995
f1_keywords:
- vbaol11.chm995
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email2Address
ms.assetid: 1656eb41-55b3-50f7-7351-b287e07bcac0
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Email2Address property (Outlook)

Returns or sets a **String** representing the email address of the second email entry for the contact. Read/write.


## Syntax

_expression_. `Email2Address`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the email address for the second email entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email2Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]