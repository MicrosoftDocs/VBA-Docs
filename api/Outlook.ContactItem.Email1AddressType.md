---
title: ContactItem.Email1AddressType property (Outlook)
keywords: vbaol11.chm992
f1_keywords:
- vbaol11.chm992
ms.prod: outlook
api_name:
- Outlook.ContactItem.Email1AddressType
ms.assetid: f498f1be-713c-7d86-28c8-fbeb6b1d3f6d
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Email1AddressType property (Outlook)

Returns or sets a  **String** representing the address type (such as EX or SMTP) of the first email entry for the contact. Read/write.


## Syntax

_expression_. `Email1AddressType`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This is a free-form text field, but it must match the actual type of an existing email transport.


## Example

This Visual Basic for Applications (VBA) example sets "SMTP" as the address type for the first email entry of a contact.


```vb
Sub SetType() 
 
 Dim myItem As Outlook.ContactItem 
 
 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 myItem.Email1Address = "someone@example.com" 
 
 myItem.Email1AddressType = "SMTP" 
 
 myItem.Display 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]