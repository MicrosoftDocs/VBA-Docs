---
title: MailItem.RetentionExpirationDate property (Outlook)
keywords: vbaol11.chm3559
f1_keywords:
- vbaol11.chm3559
ms.prod: outlook
api_name:
- Outlook.MailItem.RetentionExpirationDate
ms.assetid: 8f251c3d-8ccc-1378-ad9c-87c6e0ee7d16
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.RetentionExpirationDate property (Outlook)

Returns a  **Date** that specifies the date when the **[MailItem](Outlook.MailItem.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.


## Syntax

_expression_. `RetentionExpirationDate`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

A retention policy is enabled and disabled by an administrator for an Exchange Server on a mailbox level. This feature is available only on an Exchange mailbox with MRM version 2.0 or later enabled.

Microsoft Outlook calculates the value of this property based on the item retention start date and the retention period, if Outlook is in cache or offline mode. The Exchange Server specifies the value if Outlook is in online mode.

 In general, the retention start date for the item is determined as follows:


- Received or sent items: the retention start date is the received date.
    
- Nonrecurring calendar items: the retention start date is the appointment end date.
    
- Recurring calendar items: the retention start date is the end date of last recurrence. If there is no end date, the item never expires.
    



## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]