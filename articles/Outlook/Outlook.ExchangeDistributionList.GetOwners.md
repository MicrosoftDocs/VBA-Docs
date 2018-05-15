---
title: ExchangeDistributionList.GetOwners Method (Outlook)
keywords: vbaol11.chm2135
f1_keywords:
- vbaol11.chm2135
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.GetOwners
ms.assetid: f09f5550-b750-4e39-9644-bc98a978daa2
ms.date: 06/08/2017
---


# ExchangeDistributionList.GetOwners Method (Outlook)

Returns an  **[AddressEntries](Outlook.AddressEntries.md)** collection object that contains all the owners of the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** .


## Syntax

 _expression_ . **GetOwners**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


### Return Value

An  **AddressEntries** collection object that contains **[AddressEntry](Outlook.AddressEntry.md)** objects representing all the owners of the **ExchangeDistributionList** . Returns an **AddressEntries** object with a count of zero (0) if no owners can be found for the **ExchangeDistributionList** in the current session.


## Remarks

 **GetOwners** is an expensive operation in terms of performance if there is a slow connection to Exchange Server.


## See also


#### Concepts


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

