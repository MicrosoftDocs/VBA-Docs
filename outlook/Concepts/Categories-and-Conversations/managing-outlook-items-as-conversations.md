---
title: Managing Outlook Items as Conversations
ms.prod: outlook
ms.assetid: d91959d7-07b2-7952-8e6d-a39422d355e0
ms.date: 06/08/2019
localization_priority: Normal
---


# Managing Outlook Items as Conversations

In Microsoft Outlook, a conversation groups messages that share the same subject and belong to the same thread. In the Outlook user interface, you can expand a conversation in Conversation view to provide a visual relationship between messages, including any responses and related messages from other folders. A conversation can also include branches, such as when a message gets two or more responses and discussions grow independently from each. Since Outlook, Conversation view relates all items in the same conversation across folders and stores.

 From the programmatic perspective, items in the same conversation can be heterogeneous, belonging to one or more item types. For example, a conversation can contain **[MailItem](../../../api/Outlook.MailItem.md)** and **[TaskItem](../../../api/Outlook.TaskItem.md)** objects. Before Outlook, support for items that belong to the same conversation was limited to the **ConversationIndex** and **ConversationTopic** properties (for all item types except the **[NoteItem](../../../api/Outlook.NoteItem.md)** object). Clearing the **ConversationIndex** was limited to the **[MailItem](../../../api/Outlook.MailItem.md)**, **[PostItem](../../../api/Outlook.PostItem.md)**, and **[SharingItem](../../../api/Outlook.SharingItem.md)** objects. Since Outlook, Outlook supports the **[Conversation](../../../api/Outlook.Conversation.md)** object, which relates all items in the same conversation across folders and across stores by using the **ConversationID** property on the **Conversation** object as well as on each item of the conversation. Outlook provides a **GetConversation** method for most item types to enable you to obtain a **Conversation** object based on the item.

Conversation view is supported by stores that are POP, IMAP, PST, or Microsoft Exchange Server (at least Microsoft Exchange Server 2010, or Microsoft Exchange Server 2007 if Outlook is running in cached mode). You can call the **[IsConversationEnabled](../../../api/Outlook.Store.IsConversationEnabled.md)** property of the **[Store](../../../api/Outlook.Store.md)** object to verify whether the store supports Conversation view. You can call the **GetConversation** method to get a **Conversation** object based on an item in the conversation only if the store in which the item resides supports Conversation view.

To navigate a conversation hierarchy, you can call the **[GetChildren](../../../api/Outlook.Conversation.GetChildren.md)**, **[GetParent](../../../api/Outlook.Conversation.GetParent.md)**, and **[GetRootItems](../../../api/Outlook.Conversation.GetRootItems.md)** methods of the **Conversation** object. The **[SimpleItems](../../../api/Outlook.SimpleItems.md)** collection exists to provide easy access to items of the conversation. The order of items in the **SimpleItems** collection is the same as the order of items in the conversation. The collection is ordered by the MAPI **PidTagCreationTime** property of each item in ascending order.
To enumerate items in a conversation, you can use the **[Table](../../../api/Outlook.Table.md)** object. The rows of the table represent items of the conversation, and the columns of the table, which you can customize, represent properties for each item. To obtain conversation items by using a **Table** object, use the following procedure:

1. Obtain the object of any item in the conversation.

2. To verify that the store supports Conversation view, use the **IsConversationEnabled** property of the **Store** object that represents the store in which the item resides. You can obtain a **Conversation** object based on an item only if the item resides in a store that supports Conversation view.

3. If the store supports Conversation view, call the **GetConversation** method of that item to get the **Conversation** object.

4. Call the **[GetTable](../../../api/Outlook.Conversation.GetTable.md)** method of that **Conversation** object to get a **Table**.

5. You can now use methods that the **Table** object supports to enumerate rows that represent conversation items, and use the default columns to access default item properties (or customize columns to access other properties of the items).

You can use the **[SetAlwaysDelete](../../../api/Outlook.Conversation.SetAlwaysDelete.md)** and **[SetAlwaysMoveToFolder](../../../api/Outlook.Conversation.SetAlwaysMoveToFolder.md)** methods to always move existing conversation items, and future items that arrive in a specific conversation, to the Deleted Items folder or another folder. The moving of items is supported in the specific store where the item resides, unless the store is a non-delivery store such as a PST store. You can use the **[GetAlwaysDelete](../../../api/Outlook.Conversation.GetAlwaysDelete.md)** and **[GetAlwaysMoveToFolder](../../../api/Outlook.Conversation.GetAlwaysMoveToFolder.md)** methods to get these folders, and the **[StopAlwaysDelete](../../../api/Outlook.Conversation.StopAlwaysDelete.md)** and **[StopAlwaysMoveToFolder](../../../api/Outlook.Conversation.StopAlwaysMoveToFolder.md)** methods to stop moving existing and future conversation items to such folders.
In addition, you can apply actions to all existing and future items of a conversation.

- Call the **[SetAlwaysAssignCategories](../../../api/Outlook.Conversation.SetAlwaysAssignCategories.md)** and **[GetAlwaysAssignCategories](../../../api/Outlook.Conversation.GetAlwaysAssignCategories.md)** methods to set and get categories, respectively, for existing and future conversation items.

- Call the **[MarkAsRead](../../../api/Outlook.Conversation.MarkAsRead.md)** and **[MarkAsUnread](../../../api/Outlook.Conversation.MarkAsUnread.md)** methods to mark items as read or unread, respectively.

## See also

 [How to: Obtain and Enumerate Selected Conversations](obtain-and-enumerate-selected-conversations.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]