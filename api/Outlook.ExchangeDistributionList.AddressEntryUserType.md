---
title: ExchangeDistributionList.AddressEntryUserType property (Outlook)
keywords: vbaol11.chm2127
f1_keywords:
- vbaol11.chm2127
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.AddressEntryUserType
ms.assetid: 4b52f24d-4864-b424-a2d4-4d04d3e455ea
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.AddressEntryUserType property (Outlook)

Returns  **olExchangeDistributionListAddressEntry** which is a constant from the **[OlAddressEntryUserType](Outlook.OlAddressEntryUserType.md)** enumeration representing the user type of the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)**. Read-only.


## Syntax

_expression_. `AddressEntryUserType`

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

The  **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **AddressEntryUserType** property from the **AddressEntry** object. In the case of **ExchangeDistributionList**, **AddressEntryUserType** should always return **olExchangeDistributionListAddressEntry**.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]