---
title: MailItem.ConversationID property (Outlook)
keywords: vbaol11.chm3468
f1_keywords:
- vbaol11.chm3468
ms.prod: outlook
api_name:
- Outlook.MailItem.ConversationID
ms.assetid: 97532cd6-397b-303e-b265-7923b371bf9d
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[MailItem](Outlook.MailItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **MailItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
