---
title: MailItem.ReadReceiptRequested property (Outlook)
keywords: vbaol11.chm1340
f1_keywords:
- vbaol11.chm1340
ms.prod: outlook
api_name:
- Outlook.MailItem.ReadReceiptRequested
ms.assetid: 5b8d5283-b2fc-4b01-6ccb-b8ac6c7c617e
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ReadReceiptRequested property (Outlook)

Returns a  **Boolean** value that indicates **True** if a read receipt has been requested by the sender.


## Syntax

_expression_. `ReadReceiptRequested`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagReadReceiptRequested**. Read/write for email items that have been created but have not been sent or posted; read-only for sent email items.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]