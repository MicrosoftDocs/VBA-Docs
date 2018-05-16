---
title: ExchangeUser.ID Property (Outlook)
keywords: vbaol11.chm2067
f1_keywords:
- vbaol11.chm2067
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.ID
ms.assetid: b26ae0d3-ba96-f3ad-cd74-92ce5305e702
ms.date: 06/08/2017
---


# ExchangeUser.ID Property (Outlook)

Returns a  **String** representing the unique identifier for the **[ExchangeUser](Outlook.ExchangeUser.md)** . Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents an **ExchangeUser** object.


## Remarks

 The **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **ID** property from the **AddressEntry** object. The transport provider assigns a permanent, unique string ID when an **AddressEntry** object is created. These identifiers do not change from one session to another.


## See also


#### Concepts


[ExchangeUser Object](Outlook.ExchangeUser.md)

