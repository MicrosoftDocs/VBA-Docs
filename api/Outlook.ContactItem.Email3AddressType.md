---
title: ContactItem.Email3AddressType property (Outlook)
keywords: vbaol11.chm1000
f1_keywords:
- vbaol11.chm1000
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email3AddressType
ms.assetid: af814290-2f74-5d83-28b0-a0af055a63cc
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Email3AddressType property (Outlook)

Returns or sets a **String** representing the address type (such as EX or SMTP) of the third email entry for the contact. Read/write.


## Syntax

_expression_. `Email3AddressType`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This is a free-form text field, but it must match the actual type of an existing email transport.


## Example

This Visual Basic for Applications (VBA) example sets "SMTP" as the address type for the third email entry of a contact.


```vb
Sub SetType() 
 
 Dim myItem As ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email3Address = "someone@example.com" 
 
 myItem.Email3AddressType = "SMTP" 
 
 myItem.Display 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]