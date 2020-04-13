---
title: ExchangeUser.AddressEntryUserType property (Outlook)
keywords: vbaol11.chm2080
f1_keywords:
- vbaol11.chm2080
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.AddressEntryUserType
ms.assetid: fb5b16be-8846-7c9f-22bf-847d2cfc0a54
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.AddressEntryUserType property (Outlook)

Returns  **olExchangeUserAddressEntry** which is a constant from the **[OlAddressEntryUserType](Outlook.OlAddressEntryUserType.md)** enumeration representing the user type of the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read-only.


## Syntax

_expression_. `AddressEntryUserType`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

The **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **AddressEntryUserType** property from the **AddressEntry** object. In the case of **ExchangeUser**, **AddressEntryUserType** should always return **olExchangeUserAddressEntry**.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]