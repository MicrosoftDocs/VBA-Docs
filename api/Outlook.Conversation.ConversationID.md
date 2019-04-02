---
title: Conversation.ConversationID property (Outlook)
keywords: vbaol11.chm3467
f1_keywords:
- vbaol11.chm3467
ms.prod: outlook
api_name:
- Outlook.Conversation.ConversationID
ms.assetid: ee3cbe92-9e98-1151-1774-bd3884ab2aa3
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.ConversationID property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](Outlook.Conversation.md)** object. Read-only.


## Syntax

_expression_. `ConversationID`

_expression_ A variable that represents a [Conversation](Outlook.Conversation.md) object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId**.


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]