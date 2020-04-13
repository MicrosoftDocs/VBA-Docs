---
title: TaskRequestDeclineItem.GetConversation method (Outlook)
keywords: vbaol11.chm3502
f1_keywords:
- vbaol11.chm3502
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.GetConversation
ms.assetid: 2c6cdc44-3fb0-5cbc-dae4-a14ae2ed1fda
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.GetConversation method (Outlook)

Obtains a **[Conversation](Outlook.Conversation.md)** object that represents the conversation to which this item belongs.


## Syntax

_expression_. `GetConversation`

_expression_ A variable that represents a '[TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md)' object.


## Return value

A  **Conversation** object that represents the conversation to which this item belongs.


## Remarks

 **GetConversation** returns **Null** (**Nothing** in Visual Basic) if no conversation exists for the item. No conversation exists for an item in the following scenarios:


- The item has not been saved. An item can be saved programmatically, by user action, or by auto-save.
    
- For an item that can be sent (for example, a mail item, appointment item, or contact item), the item has not been sent.
    
- Conversations have been disabled through the Windows registry.
    
- The store does not support Conversation view (for example, Outlook is running in classic online mode against a version of Microsoft Exchange earlier than Microsoft Exchange Server 2010). Use the  **[IsConversationEnabled](Outlook.Store.IsConversationEnabled.md)** property of the **[Store](Outlook.Store.md)** object to determine whether the store supports Conversation view.
    



## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]