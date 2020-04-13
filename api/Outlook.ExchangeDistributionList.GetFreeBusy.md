---
title: ExchangeDistributionList.GetFreeBusy method (Outlook)
keywords: vbaol11.chm2122
f1_keywords:
- vbaol11.chm2122
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.GetFreeBusy
ms.assetid: b7b5ac5a-3973-a9ed-e716-50491cd5d9da
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.GetFreeBusy method (Outlook)

Returns  **Null** (**Nothing** in Visual Basic) because free-busy information is available only to individual users and not **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)** objects.


## Syntax

_expression_. `GetFreeBusy`( `_Start_` , `_MinPerChar_` , `_CompleteFormat_` )

 _expression_ An expression that returns an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

The **ExchangeDistributionList** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **GetFreeBusy** method from the **AddressEntry** object, and in the case of **ExchangeDistributionList**, regardless of the values of the parameters, this method always returns **Null**.

 This method does not return the free-busy information of individual members of an **ExchangeDistributionList**. To obtain free-busy information for a meeting request, send the request to individual users. Use the **[AddressEntry.AddressEntryUserType](Outlook.AddressEntry.AddressEntryUserType.md)** property of the **AddressEntry** object obtained from **[Recipient.AddressEntry](Outlook.Recipient.AddressEntry.md)** to determine if a **[Recipient](Outlook.Recipient.md)** represents an **ExchangeDistributionList**.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]