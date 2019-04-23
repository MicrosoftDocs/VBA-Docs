---
title: ExchangeUser.Delete method (Outlook)
keywords: vbaol11.chm2073
f1_keywords:
- vbaol11.chm2073
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Delete
ms.assetid: d11a82c4-28de-efef-5170-20f999f2bf08
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Delete method (Outlook)

Deletes the  **[ExchangeUser](Outlook.ExchangeUser.md)** object from the **[AddressEntries](Outlook.AddressEntries.md)** collection object to which it belongs.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

The  **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. An **ExchangeUser** object is an **AddressEntry** object that has **olExchangeUserAddressEntry** as the **[AddressEntry.AddressEntryUserType](Outlook.AddressEntry.AddressEntryUserType.md)**; calling **[AddressEntry.GetExchangeUser](Outlook.AddressEntry.GetExchangeUser.md)** returns the corresponding **ExchangeUser** object.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]