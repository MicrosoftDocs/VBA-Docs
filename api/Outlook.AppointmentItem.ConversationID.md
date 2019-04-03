---
title: AppointmentItem.ConversationID property (Outlook)
keywords: vbaol11.chm3469
f1_keywords:
- vbaol11.chm3469
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ConversationID
ms.assetid: 6897e23d-1d1d-f8fb-fbab-aa19242f4e7f
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.ConversationID property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object that the **[AppointmentItem](Outlook.AppointmentItem.md)** object belongs to. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.

If the  **AppointmentItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](Outlook.AppointmentItem.ConversationTopic.md)** property.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]