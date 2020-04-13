---
title: Account.AutoDiscoverXml property (Outlook)
keywords: vbaol11.chm3422
f1_keywords:
- vbaol11.chm3422
ms.prod: outlook
api_name:
- Outlook.Account.AutoDiscoverXml
ms.assetid: 201c5aba-5cff-0934-a750-b4ac0cb30860
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.AutoDiscoverXml property (Outlook)

Returns a **String** that represents information in XML retrieved from the auto-discovery service of the Microsoft Exchange Server that is associated with the account. Read-only.


## Syntax

_expression_. `AutoDiscoverXml`

_expression_ A variable that represents an '[Account](Outlook.Account.md)' object.


## Remarks

This property is similar to the  **[AutoDiscoverXml](Outlook.NameSpace.AutoDiscoverXml.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that this property applies to the account for which auto-discovery is completed and not necessarily to the primary Exchange account.

The returned string of XML contains information about various Web services (for example, availability service and unified messaging service) and available servers.

An error is returned if the account is not associated with an Exchange Server that is running Microsoft Exchange Server 2007 or later.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]