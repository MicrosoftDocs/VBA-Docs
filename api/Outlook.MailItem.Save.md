---
title: MailItem.Save method (Outlook)
keywords: vbaol11.chm1326
f1_keywords:
- vbaol11.chm1326
ms.prod: outlook
api_name:
- Outlook.MailItem.Save
ms.assetid: 7d7b5f22-4749-e908-41a7-12a4c730c695
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Save method (Outlook)

Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.


## Syntax

_expression_.**Save**

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

If a mail item is an inline reply, calling  **Save** on that mail item may fail and result in unexpected behavior.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
