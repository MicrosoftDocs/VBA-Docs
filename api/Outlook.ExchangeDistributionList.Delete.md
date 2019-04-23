---
title: ExchangeDistributionList.Delete method (Outlook)
keywords: vbaol11.chm2120
f1_keywords:
- vbaol11.chm2120
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Delete
ms.assetid: f1d14d2f-63ba-d02a-d40f-56f7d807e11e
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.Delete method (Outlook)

Deletes the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object from the **[AddressEntries](Outlook.AddressEntries.md)** collection object to which it belongs.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

The **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. 

An **ExchangeDistributionList** object is an **AddressEntry** object that has **olExchangeDistributionListAddressEntry** as the **[AddressEntry.AddressEntryUserType](Outlook.AddressEntry.AddressEntryUserType.md)**; calling **[AddressEntry.GetExchangeDistributionList](Outlook.AddressEntry.GetExchangeDistributionList.md)** returns the corresponding **ExchangeDistributionList** object.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]