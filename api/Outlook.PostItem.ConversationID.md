---
title: PostItem.ConversationID property (Outlook)
keywords: vbaol11.chm3473
f1_keywords:
- vbaol11.chm3473
ms.prod: outlook
api_name:
- Outlook.PostItem.ConversationID
ms.assetid: 102f64a0-2188-3731-eb13-95bc41da4e37
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[PostItem](Outlook.PostItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **PostItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]