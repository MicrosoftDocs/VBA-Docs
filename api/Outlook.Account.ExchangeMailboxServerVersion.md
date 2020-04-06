---
title: Account.ExchangeMailboxServerVersion property (Outlook)
keywords: vbaol11.chm3551
f1_keywords:
- vbaol11.chm3551
ms.prod: outlook
api_name:
- Outlook.Account.ExchangeMailboxServerVersion
ms.assetid: 5bfd2c63-5a87-9225-a9a8-1771fc480f21
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.ExchangeMailboxServerVersion property (Outlook)

Returns a  **String** value that represents the full version number of the Microsoft Exchange Server that hosts the account mailbox. Read-only.


## Syntax

_expression_. `ExchangeMailboxServerVersion`

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Remarks

This property is similar to the  **[ExchangeMailboxServerVersion](Outlook.NameSpace.ExchangeMailboxServerVersion.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that this property applies to the Exchange Server that hosts the account mailbox, and not necessarily to the primary Exchange account.

This property returns a string that contains the version number of the Exchange server for the account. The version number has the following four parts. 




```vb
<major version>.<minor version>.<build number>.<revision>
```

Not all parts may be present in the version number, depending on the version information that is supplied by the Exchange Server. For example, this property returns "6.5.7638" for Microsoft Exchange Server 2003 Service Pack 2.

If an Exchange mailbox is not associated with this account, this property returns an empty string.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]