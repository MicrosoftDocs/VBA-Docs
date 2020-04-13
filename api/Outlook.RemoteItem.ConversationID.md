---
title: RemoteItem.ConversationID property (Outlook)
keywords: vbaol11.chm3495
f1_keywords:
- vbaol11.chm3495
ms.prod: outlook
api_name:
- Outlook.RemoteItem.ConversationID
ms.assetid: 7cef33a7-99f8-63f6-a987-6dce94fa3120
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[RemoteItem](Outlook.RemoteItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **RemoteItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]