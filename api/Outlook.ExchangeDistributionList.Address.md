---
title: ExchangeDistributionList.Address property (Outlook)
keywords: vbaol11.chm2112
f1_keywords:
- vbaol11.chm2112
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.Address
ms.assetid: 9bfb7b5c-02ec-febc-c411-574efaa52c55
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.Address property (Outlook)

Returns or sets a **String** representing the X400 email address of the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)**. Read/write.


## Syntax

_expression_.**Address**

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

This property assumes the X400 address of the distribution list. To determine the primary Internet address, use the  **[ExchangeDistributionList.PrimarySmtpAddress](Outlook.ExchangeDistributionList.PrimarySmtpAddress.md)** property.

The **Address** property must be set before calling the **[ExchangeDistributionList.Details](Outlook.ExchangeUser.Details.md)** method.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]