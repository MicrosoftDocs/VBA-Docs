---
title: Items.SetColumns method (Outlook)
keywords: vbaol11.chm71
f1_keywords:
- vbaol11.chm71
ms.prod: outlook
api_name:
- Outlook.Items.SetColumns
ms.assetid: 90206a68-baf8-282c-5793-fee029fed452
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.SetColumns method (Outlook)

Caches certain properties for extremely fast access to those particular properties of each item in an  **[Items](Outlook.Items.md)** collection.


## Syntax

_expression_. `SetColumns`( `_Columns_` )

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Columns_|Required| **String**|A string that contains the names of the properties to cache. The property names are delimited by commas in this string.|

## Remarks

The  **SetColumns** method is useful for iterating through an **Items** collection. If you don't use this method, Microsoft Outlook must open each item to access the property. With the **SetColumns** method, Outlook only checks the properties that you have cached, and provides fast, read-only access to these properties.

After applying the  **SetColumns** method on specific properties of the collection, you cannot read other properties of that collection; properties which are not cached are returned empty. You cannot write to any of the properties of that collection either. Alternatively, if you require read-write, fast access to a set of items, use the **[Table](Outlook.Table.md)** object.

 **SetColumns** cannot be used, and will cause an error, with any property that returns an object. It cannot be used with the following properties:



| **AutoResolvedWinner**| **InternetCodePage**|
| **Body**| **MeetingWorkspaceURL**|
| **BodyFormat**| **[MemberCount](Outlook.DistListItem.MemberCount.md)**|
| **Categories**| **ReceivedByEntryID**|
| **[Children](Outlook.ContactItem.Children.md)**| **ReceivedOnBehalfOfEntryID**|
| **Class**| **[RecurrenceState](Outlook.AppointmentItem.RecurrenceState.md)**|
| **Companies**| **ReplyRecipients**|
| **[DLName](Outlook.DistListItem.DLName.md)**| **[ResponseState](Outlook.TaskItem.ResponseState.md)**|
| **DownloadState**| **Saved**|
| **EntryID**| **Sent**|
| **HTMLBody**| **Submitted**|
| **IsConflict**| **[VotingOptions](Outlook.MailItem.VotingOptions.md)**|

The  **ConversationIndex** property cannot be cached using the **SetColumns** method. However, this property will not result in an error like the other properties listed above.


## Example

The following Visual Basic for Applications (VBA) example uses the  **[Items](Outlook.Items.md)** collection to get the items in default Tasks folder, caches the **[Subject](Outlook.MailItem.Subject.md)** and **[DueDate](Outlook.TaskItem.DueDate.md)** properties and then displays the subject and due dates each in turn.


```vb
Sub SortByDueDate() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 Dim myItems As Outlook.Items 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderTasks) 
 
 Set myItems = myFolder.Items 
 
 myItems.SetColumns ("Subject, DueDate") 
 
 For Each myItem In myItems 
 
 MsgBox myItem.Subject & " " & myItem.DueDate 
 
 Next myItem 
 
End Sub
```


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]