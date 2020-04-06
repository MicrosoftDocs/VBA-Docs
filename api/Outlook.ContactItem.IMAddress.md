---
title: ContactItem.IMAddress property (Outlook)
keywords: vbaol11.chm1085
f1_keywords:
- vbaol11.chm1085
ms.prod: outlook
api_name:
- Outlook.ContactItem.IMAddress
ms.assetid: d7f916b0-aa5b-872d-0928-bbab5000ac75
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.IMAddress property (Outlook)

Returns or sets a  **String** that represents a contact's Microsoft Instant Messenger address. Read/write.


## Syntax

_expression_. `IMAddress`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

Unlike the  **[Recipients](Outlook.MailItem.Recipients.md)** or **[To](Outlook.MailItem.To.md)** properties, there is no way to verify that the **IMAddress** property contains a valid address.


## Example

The following example creates a new contact and prompts the user to enter an Instant Messenger address for the contact.


```vb
Sub SetImAddress() 
 
 'Sets a new IM Address 
 
 Dim objNewContact As ContactItem 
 
 
 
 Set objNewContact = Application.CreateItem(olContactItem) 
 
 objNewContact.IMAddress = _ 
 
 InputBox("Enter the new contact's Microsoft Instant Messenger address") 
 
 objNewContact.Save 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]