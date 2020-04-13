---
title: Conversation object (Outlook)
keywords: vbaol11.chm3388
f1_keywords:
- vbaol11.chm3388
ms.prod: outlook
api_name:
- Outlook.Conversation
ms.assetid: 2705d38a-ebc0-e5a7-208b-ffe1f5446b1b
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation object (Outlook)

Represents a conversation that includes one or more items stored in one or more folders and stores.


## Remarks

The **Conversation** object is an abstract, aggregated object. Although a conversation can include items of different types, the **Conversation** object does not correspond to a particular underlying MAPI **IMessage** object.

A conversation represents one or more items in one or more folders and stores. If you move an item in a conversation to the  **Deleted Items** folder and subsequently enumerate the conversation by using the **[GetChildren](Outlook.Conversation.GetChildren.md)**, **[GetRootItems](Outlook.Conversation.GetRootItems.md)**, or **[GetTable](Outlook.Conversation.GetTable.md)** method, the item will not be included in the returned object.

To obtain a **Conversation** object for an existing conversation, use the **GetConversation** method of the item.

There are actions that you can apply to items in a conversation by calling the  **[SetAlwaysAssignCategories](Outlook.Conversation.SetAlwaysAssignCategories.md)**, **[SetAlwaysDelete](Outlook.Conversation.SetAlwaysDelete.md)**, or **[SetAlwaysMoveToFolder](Outlook.Conversation.SetAlwaysMoveToFolder.md)** method. Each of these actions is applied to all items in the conversation automatically when the method is called; the action is also applied to future items in the conversation as long as the action is still applicable to the conversation. There is no explicit save method on the **Conversation** object.

Also, when you apply an action to items in a conversation, the corresponding event occurs. For example, the  **[ItemChange](Outlook.Items.ItemChange.md)** event of the **[Items](Outlook.Items.md)** object occurs when you call **SetAlwaysAssignCategories**, and the **[BeforeItemMove](Outlook.Folder.BeforeItemMove.md)** event of the **[Folder](Outlook.Folder.md)** object occurs when you call **SetAlwaysMoveToFolder**.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example assumes that the selected item in the explorer window is a mail item. The code example gets the conversation that the selected mail item is associated with, and enumerates each item in that conversation, displaying the subject of the item. The  `DemoConversation` method calls the **GetConversation** method of the selected mail item to get the associated **Conversation** object. `DemoConversation` then calls the **[GetTable](Outlook.Conversation.GetTable.md)** and **[GetRootItems](Outlook.Conversation.GetRootItems.md)** methods of the **Conversation** object to get a **[Table](Outlook.Table.md)** object and **[SimpleItems](Outlook.SimpleItems.md)** collection, respectively. `DemoConversation` calls the recurrent method `EnumerateConversation` to enumerate and display the subject of each item in that conversation.




```cs
void DemoConversation() 
{ 
 object selectedItem = 
 Application.ActiveExplorer().Selection[1]; 
 // This example uses only 
 // MailItem. Other item types such as 
 // MeetingItem and PostItem can participate 
 // in the conversation. 
 if (selectedItem is Outlook.MailItem) 
 { 
 // Cast selectedItem to MailItem. 
 Outlook.MailItem mailItem = 
 selectedItem as Outlook.MailItem; 
 // Determine the store of the mail item. 
 Outlook.Folder folder = mailItem.Parent 
 as Outlook.Folder; 
 Outlook.Store store = folder.Store; 
 if (store.IsConversationEnabled == true) 
 { 
 // Obtain a Conversation object. 
 Outlook.Conversation conv = 
 mailItem.GetConversation(); 
 // Check for null Conversation. 
 if (conv != null) 
 { 
 // Obtain Table that contains rows 
 // for each item in the conversation. 
 Outlook.Table table = conv.GetTable(); 
 Debug.WriteLine("Conversation Items Count: " + 
 table.GetRowCount().ToString()); 
 Debug.WriteLine("Conversation Items from Table:"); 
 while (!table.EndOfTable) 
 { 
 Outlook.Row nextRow = table.GetNextRow(); 
 Debug.WriteLine(nextRow["Subject"] 
 + " Modified: " 
 + nextRow["LastModificationTime"]); 
 } 
 Debug.WriteLine("Conversation Items from Root:"); 
 // Obtain root items and enumerate the conversation. 
 Outlook.SimpleItems simpleItems 
 = conv.GetRootItems(); 
 foreach (object item in simpleItems) 
 { 
 // In this example, only enumerate MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (item is Outlook.MailItem) 
 { 
 Outlook.MailItem mail = item 
 as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mail.Parent as Outlook.Folder; 
 string msg = mail.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Call EnumerateConversation 
 // to access child nodes of root items. 
 EnumerateConversation(item, conv); 
 } 
 } 
 } 
 } 
} 
 
 
void EnumerateConversation(object item, 
 Outlook.Conversation conversation) 
{ 
 Outlook.SimpleItems items = 
 conversation.GetChildren(item); 
 if (items.Count > 0) 
 { 
 foreach (object myItem in items) 
 { 
 // In this example, only enumerate MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (myItem is Outlook.MailItem) 
 { 
 Outlook.MailItem mailItem = 
 myItem as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mailItem.Parent as Outlook.Folder; 
 string msg = mailItem.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Continue recursion. 
 EnumerateConversation(myItem, conversation); 
 } 
 } 
} 
 

```


## Methods



|Name|
|:-----|
|[ClearAlwaysAssignCategories](Outlook.Conversation.ClearAlwaysAssignCategories.md)|
|[GetAlwaysAssignCategories](Outlook.Conversation.GetAlwaysAssignCategories.md)|
|[GetAlwaysDelete](Outlook.Conversation.GetAlwaysDelete.md)|
|[GetAlwaysMoveToFolder](Outlook.Conversation.GetAlwaysMoveToFolder.md)|
|[GetChildren](Outlook.Conversation.GetChildren.md)|
|[GetParent](Outlook.Conversation.GetParent.md)|
|[GetRootItems](Outlook.Conversation.GetRootItems.md)|
|[GetTable](Outlook.Conversation.GetTable.md)|
|[MarkAsRead](Outlook.Conversation.MarkAsRead.md)|
|[MarkAsUnread](Outlook.Conversation.MarkAsUnread.md)|
|[SetAlwaysAssignCategories](Outlook.Conversation.SetAlwaysAssignCategories.md)|
|[SetAlwaysDelete](Outlook.Conversation.SetAlwaysDelete.md)|
|[SetAlwaysMoveToFolder](Outlook.Conversation.SetAlwaysMoveToFolder.md)|
|[StopAlwaysDelete](Outlook.Conversation.StopAlwaysDelete.md)|
|[StopAlwaysMoveToFolder](Outlook.Conversation.StopAlwaysMoveToFolder.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Conversation.Application.md)|
|[Class](Outlook.Conversation.Class.md)|
|[ConversationID](Outlook.Conversation.ConversationID.md)|
|[Parent](Outlook.Conversation.Parent.md)|
|[Session](Outlook.Conversation.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Conversation Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]