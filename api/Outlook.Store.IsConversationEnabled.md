---
title: Store.IsConversationEnabled property (Outlook)
keywords: vbaol11.chm3518
f1_keywords:
- vbaol11.chm3518
ms.prod: outlook
api_name:
- Outlook.Store.IsConversationEnabled
ms.assetid: ce333881-a5f3-2115-0ae4-296d15c4bead
ms.date: 06/27/2019
localization_priority: Normal
---


# Store.IsConversationEnabled property (Outlook)

Returns a **Boolean** value that is **True** if the store supports Conversation view. Read-only.


## Syntax

_expression_.**IsConversationEnabled**

_expression_ A variable that represents a **[Store](Outlook.Store.md)** object.


## Remarks

A store supports Conversation view if the store is a POP, IMAP, or PST store, or if it runs a version of Microsoft Exchange Server that is at least Microsoft Exchange Server 2010. A store also supports Conversation view if the store is running Microsoft Exchange Server 2007, the version of Outlook is at least Outlook 2010, and Outlook is running in cached mode.

If a store supports conversations, calling the **GetConversation** method of an item in the store returns a **[Conversation](Outlook.Conversation.md)** object for the item. If the store does not support conversations, **GetConversation** returns **Null** (**Nothing** in Visual Basic) for items in the store.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
