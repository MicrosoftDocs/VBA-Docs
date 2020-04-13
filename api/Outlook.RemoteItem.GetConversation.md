---
title: RemoteItem.GetConversation method (Outlook)
keywords: vbaol11.chm3494
f1_keywords:
- vbaol11.chm3494
ms.prod: outlook
api_name:
- Outlook.RemoteItem.GetConversation
ms.assetid: 0da0ef2e-c3d2-5a5f-c031-4226719f87f7
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.GetConversation method (Outlook)

Obtains a **[Conversation](Outlook.Conversation.md)** object that represents the conversation to which this item belongs.


## Syntax

_expression_. `GetConversation`

_expression_ A variable that represents a '[RemoteItem](Outlook.RemoteItem.md)' object.


## Return value

A **Conversation** object that represents the conversation to which this item belongs.


## Remarks

 **GetConversation** returns **Null** (**Nothing** in Visual Basic) if no conversation exists for the item. No conversation exists for an item in the following scenarios:


- The item has not been saved. An item can be saved programmatically, by user action, or by auto-save.
    
- For an item that can be sent (for example, a mail item, appointment item, or contact item), the item has not been sent.
    
- Conversations have been disabled through the Windows registry.
    
- The store does not support Conversation view (for example, Outlook is running in classic online mode against a version of Microsoft Exchange earlier than Microsoft Exchange Server 2010). Use the  **[IsConversationEnabled](Outlook.Store.IsConversationEnabled.md)** property of the **[Store](Outlook.Store.md)** object to determine whether the store supports Conversation view.
    



## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]