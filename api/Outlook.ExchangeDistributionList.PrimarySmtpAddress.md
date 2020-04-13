---
title: ExchangeDistributionList.PrimarySmtpAddress property (Outlook)
keywords: vbaol11.chm2134
f1_keywords:
- vbaol11.chm2134
ms.prod: outlook
api_name:
- Outlook.ExchangeDistributionList.PrimarySmtpAddress
ms.assetid: f64bbc29-14c4-be68-402a-16d9ac34a727
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeDistributionList.PrimarySmtpAddress property (Outlook)

Returns a **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)**. Read-only.


## Syntax

_expression_. `PrimarySmtpAddress`

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagEmailAddress**.

Returns an empty string if this property has not been implemented or does not exist for the  **ExchangeDistributionList** object.


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)



[How to: Obtain the Email Address of a Recipient](../outlook/Concepts/Address-Book/obtain-the-e-mail-address-of-a-recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]