---
title: ExchangeUser.GetContact method (Outlook)
keywords: vbaol11.chm2078
f1_keywords:
- vbaol11.chm2078
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetContact
ms.assetid: 443fb23a-cd26-e385-bd9d-e978aec56458
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.GetContact method (Outlook)

Returns  **Null** (**Nothing** in Visual Basic) because the **[ExchangeUser](Outlook.ExchangeUser.md)** object does not correspond to a contact in a Contacts Address Book.


## Syntax

_expression_. `GetContact`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Return value

 **Null** (**Nothing** in Visual Basic) because the **ExchangeUser** object does not correspond to a contact in a Contacts Address Book.


## Remarks

The  **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **GetContact** method from the **AddressEntry** object, and in the case of **ExchangeUser**, this method always returns **Null**.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]