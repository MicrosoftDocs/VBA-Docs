---
title: ExchangeDistributionList.GetContact method (Outlook)
keywords: vbaol11.chm2125
f1_keywords:
- vbaol11.chm2125
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.GetContact
ms.assetid: ed3cf7ab-32b9-6dad-66d5-a4cd2ad9a9f4
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.GetContact method (Outlook)

Returns  **Null** (**Nothing** in Visual Basic) because the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object does not correspond to a contact in a Contacts Address Book.


## Syntax

_expression_. `GetContact`

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Return value

 **Null** (**Nothing** in Visual Basic) because the **ExchangeDistributionList** object does not correspond to a contact in a Contacts Address Book.


## Remarks

The **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **GetContact** method from the **AddressEntry** object, and in the case of **ExchangeDistributionList**, this method always returns **Null**.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]