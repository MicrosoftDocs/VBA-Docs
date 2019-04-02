---
title: AddressEntry.GetContact method (Outlook)
keywords: vbaol11.chm2055
f1_keywords:
- vbaol11.chm2055
ms.prod: outlook
api_name:
- Outlook.AddressEntry.GetContact
ms.assetid: 2364f180-475d-aff1-01e8-30a54e870404
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntry.GetContact method (Outlook)

Returns a  **[ContactItem](Outlook.ContactItem.md)** object that represents the **[AddressEntry](Outlook.AddressEntry.md)**, if the **AddressEntry** corresponds to a contact in an Outlook Contacts Address Book (CAB).


## Syntax

_expression_. `GetContact`

_expression_ A variable that represents an [AddressEntry](Outlook.AddressEntry.md) object.


## Return value

A  **ContactItem** object that corresponds to the **AddressEntry**. Returns **Null** (**Nothing** in Visual Basic) if the **AddressEntry** object does not correspond to a contact in a Contacts Address Book.


## See also


[AddressEntry Object](Outlook.AddressEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]