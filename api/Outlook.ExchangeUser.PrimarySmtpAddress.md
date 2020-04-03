---
title: ExchangeUser.PrimarySmtpAddress property (Outlook)
keywords: vbaol11.chm2098
f1_keywords:
- vbaol11.chm2098
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.PrimarySmtpAddress
ms.assetid: 2dda21da-44a2-fbfe-babc-58646c76689d
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.PrimarySmtpAddress property (Outlook)

Returns a **String** representing the primary Simple Mail Transfer Protocol (SMTP) address for the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read-only.


## Syntax

_expression_. `PrimarySmtpAddress`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property corresponds to the MAPI property,  **PidTagEmailAddress**.

 Returns an empty string if this property has not been implemented or does not exist for the **ExchangeUser** object.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)



[How to: Obtain the Email Address of a Recipient](../outlook/Concepts/Address-Book/obtain-the-e-mail-address-of-a-recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]