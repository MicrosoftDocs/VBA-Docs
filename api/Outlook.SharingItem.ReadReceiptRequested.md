---
title: SharingItem.ReadReceiptRequested property (Outlook)
keywords: vbaol11.chm643
f1_keywords:
- vbaol11.chm643
ms.prod: outlook
api_name:
- Outlook.SharingItem.ReadReceiptRequested
ms.assetid: fa8f3b1c-77a6-1620-f0dd-7cf0bd6f64a3
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.ReadReceiptRequested property (Outlook)

Returns a **Boolean** value that indicates **True** if a read receipt has been requested by the sender.


## Syntax

_expression_. `ReadReceiptRequested`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagReadReceiptRequested**. This property is read/write for **[SharingItem](Outlook.SharingItem.md)** objects that have been created but have not been sent or posted; it is read-only for sent **SharingItem** objects.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]