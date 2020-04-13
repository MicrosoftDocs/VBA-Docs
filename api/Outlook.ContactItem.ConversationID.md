---
title: ContactItem.ConversationID property (Outlook)
keywords: vbaol11.chm3470
f1_keywords:
- vbaol11.chm3470
ms.prod: outlook
api_name:
- Outlook.ContactItem.ConversationID
ms.assetid: 13a4e7cf-66b3-fba6-b179-68eaf1de8db6
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.ConversationID property (Outlook)

Returns a **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[ContactItem](Outlook.ContactItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **ContactItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]