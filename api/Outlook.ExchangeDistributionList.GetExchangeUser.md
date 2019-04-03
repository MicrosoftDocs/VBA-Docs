---
title: ExchangeDistributionList.GetExchangeUser method (Outlook)
keywords: vbaol11.chm2126
f1_keywords:
- vbaol11.chm2126
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.GetExchangeUser
ms.assetid: a5ce23e5-76cb-ac86-b8c7-a4e63eda560d
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.GetExchangeUser method (Outlook)

Returns  **Null** (**Nothing** in Visual Basic) because the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** object does not correspond to an **[ExchangeUser](Outlook.ExchangeUser.md)** object.


## Syntax

_expression_. `GetExchangeUser`

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Return value

 **Null** (**Nothing** in Visual Basic) because the **ExchangeDistributionList** object does not correspond to an **ExchangeUser** object.


## Remarks

The  **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **GetExchangeUser** method from the **AddressEntry** object, and in the case of **ExchangeDistributionList**, this method always returns **Null**.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]