---
title: Store.ExchangeStoreType property (Outlook)
keywords: vbaol11.chm802
f1_keywords:
- vbaol11.chm802
ms.prod: outlook
api_name:
- Outlook.Store.ExchangeStoreType
ms.assetid: ca6002bd-444d-a111-adca-6f8fafc37ea1
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.ExchangeStoreType property (Outlook)

Returns a constant in the  **[OlExchangeStoreType](Outlook.OlExchangeStoreType.md)** enumeration that indicates the type of an Exchange store. Read-only.


## Syntax

_expression_. `ExchangeStoreType`

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

The  **ExchangeStoreType** property distinguishes among different Exchange store types, such as primary Exchange mailbox, Exchange mailbox, Public Folder store, or non-Exchange store. This property does not distinguish among every type of store including Hotmail, HTTP, IMAP, and so forth. Use **[Account.AccountType](Outlook.Account.AccountType.md)** for the type of server associated with an email account, such as Exchange, HTTP, IMAP, or POP3.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]