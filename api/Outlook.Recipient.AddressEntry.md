---
title: Recipient.AddressEntry property (Outlook)
keywords: vbaol11.chm2345
f1_keywords:
- vbaol11.chm2345
ms.prod: outlook
api_name:
- Outlook.Recipient.AddressEntry
ms.assetid: 3b2b524e-4dd5-9ff4-98cc-811746ea0453
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.AddressEntry property (Outlook)

Returns the  **[AddressEntry](Outlook.AddressEntry.md)** object corresponding to the resolved recipient. Read/write.


## Syntax

_expression_. `AddressEntry`

_expression_ A variable that represents a [Recipient](Outlook.Recipient.md) object.


## Remarks

Accessing the  **AddressEntry** property forces resolution of an unresolved recipient name. If the name cannot be resolved, an error is returned. If the recipient is resolved, the **[Resolved](Outlook.Recipient.Resolved.md)** property is **True**.


## See also


[Recipient Object](Outlook.Recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]