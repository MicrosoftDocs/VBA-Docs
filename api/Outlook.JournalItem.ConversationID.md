---
title: JournalItem.ConversationID property (Outlook)
keywords: vbaol11.chm3471
f1_keywords:
- vbaol11.chm3471
ms.prod: outlook
api_name:
- Outlook.JournalItem.ConversationID
ms.assetid: 05e9ccd7-af1a-a2e9-2b86-9687e0bf24c6
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[JournalItem](Outlook.JournalItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **JournalItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]