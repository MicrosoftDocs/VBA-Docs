---
title: Account.ExchangeConnectionMode property (Outlook)
keywords: vbaol11.chm3424
f1_keywords:
- vbaol11.chm3424
ms.prod: outlook
api_name:
- Outlook.Account.ExchangeConnectionMode
ms.assetid: 40fee809-48ab-5788-819a-c61b6eb782a5
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.ExchangeConnectionMode property (Outlook)

Returns an **[OlExchangeConnectionMode](Outlook.OlExchangeConnectionMode.md)** constant that indicates the current connection mode for the Microsoft Exchange Server that hosts the account mailbox. Read-only


## Syntax

_expression_. `ExchangeConnectionMode`

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Remarks

This property is similar to the  **[ExchangeConnectionMode](Outlook.NameSpace.ExchangeConnectionMode.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that this property applies to the Exchange Server that hosts the account mailbox, and not necessarily to the primary Exchange account.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]