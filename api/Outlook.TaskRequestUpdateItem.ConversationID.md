---
title: TaskRequestUpdateItem.ConversationID property (Outlook)
keywords: vbaol11.chm3506
f1_keywords:
- vbaol11.chm3506
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.ConversationID
ms.assetid: e70b6b6d-c6ba-4097-ab83-b1d826b1a6d5
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.ConversationID property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **TaskRequestUpdateItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]