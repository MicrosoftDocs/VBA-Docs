---
title: Account.AutoDiscoverConnectionMode property (Outlook)
keywords: vbaol11.chm3436
f1_keywords:
- vbaol11.chm3436
ms.prod: outlook
api_name:
- Outlook.Account.AutoDiscoverConnectionMode
ms.assetid: d9089143-caff-6e08-cc7d-f8659384d36e
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.AutoDiscoverConnectionMode property (Outlook)

Returns an  **[OlAutoDiscoverConnectionMode](Outlook.OlAutoDiscoverConnectionMode.md)** constant that specifies the type of connection to use for the auto-discovery service of the Microsoft Exchange server that hosts the account mailbox. Read-only.


## Syntax

_expression_. `AutoDiscoverConnectionMode`

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Remarks

This property is similar to the  **[AutoDiscoverConnectionMode](Outlook.NameSpace.AutoDiscoverConnectionMode.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that this property applies to the account for which auto-discovery is completed and not necessarily to the primary Exchange account.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]