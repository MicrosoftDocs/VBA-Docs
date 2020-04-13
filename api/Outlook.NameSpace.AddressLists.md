---
title: NameSpace.AddressLists property (Outlook)
keywords: vbaol11.chm759
f1_keywords:
- vbaol11.chm759
ms.prod: outlook
api_name:
- Outlook.NameSpace.AddressLists
ms.assetid: 68b236db-f964-6f7f-6246-e79c6ada19e9
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.AddressLists property (Outlook)

Returns an **[AddressLists](Outlook.AddressLists.md)** collection representing a collection of the address lists available for this session. Read-only.


## Syntax

_expression_. `AddressLists`

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Remarks

The **AddressLists** collection represents the root of the address book hierarchy for the current session. A particular **[AddressList](Outlook.AddressList.md)** object represents one of the available address books. The type of access you obtain depends on the access permissions granted to you by each individual address book provider.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]