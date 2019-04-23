---
title: MailItem.To property (Outlook)
keywords: vbaol11.chm1362
f1_keywords:
- vbaol11.chm1362
ms.prod: outlook
api_name:
- Outlook.MailItem.To
ms.assetid: 036dc0b7-1ac7-3884-8d3e-e2f2f1e66ff5
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.To property (Outlook)

Returns or sets a semicolon-delimited  **String** list of display names for the **To** recipients for the Outlook item. Read/write.


## Syntax

_expression_. `To`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property contains the display names only. The  **To** property corresponds to the MAPI property **PidTagDisplayTo**. The **[Recipients](Outlook.Recipients.md)** collection should be used to modify this property.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
