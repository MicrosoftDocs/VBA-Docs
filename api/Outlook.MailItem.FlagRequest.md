---
title: MailItem.FlagRequest property (Outlook)
keywords: vbaol11.chm1336
f1_keywords:
- vbaol11.chm1336
ms.prod: outlook
api_name:
- Outlook.MailItem.FlagRequest
ms.assetid: 13c04300-ec2a-4ee5-d7b1-eff9f61b71c4
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.FlagRequest property (Outlook)

Returns or sets a **String** that indicates the requested action for a mail item. Read/write.


## Syntax

_expression_. `FlagRequest`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

By default, a mail item is not marked with any flag and the default value for this property is the empty string. You can set the value of  **FlagRequest** either through the user interface or programmatically. When you mark a mail item with a flag through the user interface, the default value of **FlagRequest** is "Follow up".


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]