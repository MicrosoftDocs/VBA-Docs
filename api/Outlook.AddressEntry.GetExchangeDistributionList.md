---
title: AddressEntry.GetExchangeDistributionList method (Outlook)
keywords: vbaol11.chm2058
f1_keywords:
- vbaol11.chm2058
ms.prod: outlook
api_name:
- Outlook.AddressEntry.GetExchangeDistributionList
ms.assetid: 060ac302-b916-d85d-5ba8-c682894129e2
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntry.GetExchangeDistributionList method (Outlook)

Returns an **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object that represents the **[AddressEntry](Outlook.AddressEntry.md)** if the **AddressEntry** belongs to an Exchange **[AddressList](Outlook.AddressList.md)** object such as the Global Address List (GAL) and corresponds to an Exchange distribution list.


## Syntax

_expression_. `GetExchangeDistributionList`

_expression_ A variable that represents an [AddressEntry](Outlook.AddressEntry.md) object.


## Return value

An **ExchangeDistributionList** object that represents the **AddressEntry**. Returns **Null** (**Nothing** in Visual Basic) if the **AddressEntry** object does not correspond to an Exchange distribution list.


## Remarks

 You have to be connected to the Exchange server to use this method.


## See also


[AddressEntry Object](Outlook.AddressEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]