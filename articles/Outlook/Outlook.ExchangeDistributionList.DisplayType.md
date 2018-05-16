---
title: ExchangeDistributionList.DisplayType Property (Outlook)
keywords: vbaol11.chm2113
f1_keywords:
- vbaol11.chm2113
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.DisplayType
ms.assetid: e75c09e0-6acc-92cc-51a2-d43c13dd85c4
ms.date: 06/08/2017
---


# ExchangeDistributionList.DisplayType Property (Outlook)

Returns  **olDistList** which is a constant from the **[OlDisplayType](Outlook.OlDisplayType.md)** enumeration representing the nature of the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** . Read-only.


## Syntax

 _expression_ . **DisplayType**

 _expression_ A variable that represents an **ExchangeDistributionList** object.


## Remarks

This property corresponds to the MAPI property  **PidTagDisplayType** .

The  **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **DisplayType** property from the **AddressEntry** object. In the case of **ExchangeDistributionList** , **DisplayType** should always return **olDistList** .


## See also


#### Concepts


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

