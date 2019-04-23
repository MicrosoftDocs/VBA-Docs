---
title: TaskRequestAcceptItem.ConversationID property (Outlook)
keywords: vbaol11.chm3501
f1_keywords:
- vbaol11.chm3501
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.ConversationID
ms.assetid: 0cd2c84f-0332-73aa-097e-5944bf6258c8
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.ConversationID property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **TaskRequestAcceptItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]