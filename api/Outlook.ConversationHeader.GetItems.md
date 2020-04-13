---
title: ConversationHeader.GetItems method (Outlook)
keywords: vbaol11.chm3544
f1_keywords:
- vbaol11.chm3544
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.GetItems
ms.assetid: 018fab26-3cdc-cd39-4a16-fb2a26ae237f
ms.date: 06/08/2017
localization_priority: Normal
---


# ConversationHeader.GetItems method (Outlook)

Obtains a **[SimpleItems](Outlook.SimpleItems.md)** collection that contains all of the items in the conversation that reside in the same folder as the selected conversation header.


## Syntax

_expression_. `GetItems`

_expression_ A variable that represents a '[ConversationHeader](Outlook.ConversationHeader.md)' object.


## Return value

A **SimpleItems** collection of items that belong to the same conversation and reside in the same folder as the conversation header.


## Remarks

The **SimpleItems** collection only contains conversation items in the folder that contains the conversation header. The **SimpleItems** collection does not return cross-folder conversation items. If you must access cross-folder content, use the **[Conversation](Outlook.Conversation.md)** object.

If no conversation items exist in the same folder as the conversation header,  **GetItems** returns a **SimpleItems** collection with the **[SimpleItems.Count](Outlook.SimpleItems.Count.md)** property equal to 0.


## See also


[ConversationHeader Object](Outlook.ConversationHeader.md)



[How to: Obtain and Enumerate Selected Conversations](../outlook/Concepts/Categories-and-Conversations/obtain-and-enumerate-selected-conversations.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]