---
title: ExchangeUser.Address property (Outlook)
keywords: vbaol11.chm2065
f1_keywords:
- vbaol11.chm2065
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Address
ms.assetid: b3a36b16-e652-9e3f-86fd-7cea0c72d78c
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Address property (Outlook)

Returns or sets a  **String** representing the X400 email address of the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read/write.


## Syntax

_expression_.**Address**

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property assumes the X400 address of the user. To determine the primary Internet address, use the  **[ExchangeUser.PrimarySmtpAddress](Outlook.ExchangeUser.PrimarySmtpAddress.md)** property.

The  **Address** property must be set before calling the **[ExchangeUser.Details](Outlook.ExchangeUser.Details.md)** method.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]