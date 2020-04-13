---
title: TaskRequestItem.ConversationID property (Outlook)
keywords: vbaol11.chm3499
f1_keywords:
- vbaol11.chm3499
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.ConversationID
ms.assetid: 9518e9aa-a20e-69fa-b173-e90f45d28331
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[TaskRequestItem](Outlook.TaskRequestItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **TaskRequestItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]