---
title: ContactItem.Email1Address property (Outlook)
keywords: vbaol11.chm991
f1_keywords:
- vbaol11.chm991
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email1Address
ms.assetid: 0bd407bc-21a9-16e6-709d-383cb79b4d6e
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Email1Address property (Outlook)

Returns or sets a  **String** representing the email address of the first email entry for the contact. Read/write.


## Syntax

_expression_. `Email1Address`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Example

This Visual Basic for Applications (VBA) example sets "someone@example.com" as the email address for the first email entry of a contact.


```vb
Sub CreatePeerContact() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email1Address = "someone@example.com" 
 
 myItem.Display 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]