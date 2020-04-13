---
title: ExchangeUser.GetExchangeDistributionList method (Outlook)
keywords: vbaol11.chm2081
f1_keywords:
- vbaol11.chm2081
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetExchangeDistributionList
ms.assetid: 4ebc0448-97a9-ca5c-35f0-ef852de27324
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.GetExchangeDistributionList method (Outlook)

Returns  **Null** (**Nothing** in Visual Basic) because the **[ExchangeUser](Outlook.ExchangeUser.md)** object does not correspond to an **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object.


## Syntax

_expression_. `GetExchangeDistributionList`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Return value

 **Null** (**Nothing** in Visual Basic) because the **ExchangeUser** object does not correspond to an **ExchangeDistributionList** object.


## Remarks

The **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **GetExchangeDistributionList** method from the **AddressEntry** object, and in the case of **ExchangeUser**, this method always returns **Null**.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]