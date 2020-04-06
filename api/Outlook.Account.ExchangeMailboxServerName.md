---
title: Account.ExchangeMailboxServerName property (Outlook)
keywords: vbaol11.chm3550
f1_keywords:
- vbaol11.chm3550
ms.prod: outlook
api_name:
- Outlook.Account.ExchangeMailboxServerName
ms.assetid: f75354c9-3374-140f-63a6-ca04ce6101cb
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.ExchangeMailboxServerName property (Outlook)

Returns a  **String** value that represents the name of the Microsoft Exchange Server that hosts the account mailbox. Read-only.


## Syntax

_expression_. `ExchangeMailboxServerName`

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Remarks

This property is similar to the  **[ExchangeMailboxServerName](Outlook.NameSpace.ExchangeMailboxServerName.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that this property applies to the Exchange Server that hosts the account mailbox, and not necessarily to the primary Exchange account.

If an Exchange mailbox is not associated with this account, this property returns an empty string.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]