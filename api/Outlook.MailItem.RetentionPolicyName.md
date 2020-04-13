---
title: MailItem.RetentionPolicyName property (Outlook)
keywords: vbaol11.chm3558
f1_keywords:
- vbaol11.chm3558
ms.prod: outlook
api_name:
- Outlook.MailItem.RetentionPolicyName
ms.assetid: 27e2c3da-ff1a-c261-72cc-b915d89e1019
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.RetentionPolicyName property (Outlook)

Returns a **String** that specifies the name of the retention policy. Read-only.


## Syntax

_expression_. `RetentionPolicyName`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

Retention is enabled and disabled by an administrator for an Exchange Server on a mailbox level. The feature is available only on an Exchange mailbox with Messaging Records Management (MRM), version 2.0 or later enabled. An example of a retention policy name is "Define time interval for expiration Quick Searches".


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]