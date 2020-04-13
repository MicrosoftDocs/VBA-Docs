---
title: SharingItem.ReceivedByEntryID property (Outlook)
keywords: vbaol11.chm644
f1_keywords:
- vbaol11.chm644
ms.prod: outlook
api_name:
- Outlook.SharingItem.ReceivedByEntryID
ms.assetid: 8255da4d-8312-3ed5-b216-5ddc9298c505
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.ReceivedByEntryID property (Outlook)

Returns a **String** representing the **[EntryID](Outlook.Recipient.EntryID.md)** for the true recipient as set by the transport provider delivering the **[SharingItem](Outlook.SharingItem.md)**. Read-only.


## Syntax

_expression_. `ReceivedByEntryID`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagReceivedByEntryId**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]