---
title: MeetingItem.ConversationID property (Outlook)
keywords: vbaol11.chm3472
f1_keywords:
- vbaol11.chm3472
ms.prod: outlook
api_name:
- Outlook.MeetingItem.ConversationID
ms.assetid: 67a28933-1f89-8f1d-9217-bacd61aa85b9
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.ConversationID property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[MeetingItem](Outlook.MeetingItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **MeetingItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]