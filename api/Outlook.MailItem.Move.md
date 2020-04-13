---
title: MailItem.Move method (Outlook)
keywords: vbaol11.chm1324
f1_keywords:
- vbaol11.chm1324
ms.prod: outlook
api_name:
- Outlook.MailItem.Move
ms.assetid: 08a0fa20-b891-393a-00fa-5a8fb5405cf6
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Move method (Outlook)

Moves a Microsoft Outlook item to a new folder.


## Syntax

_expression_. `Move`( `_DestFldr_` )

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestFldr_|Required| **[Folder](Outlook.Folder.md)**|An expression that returns a **Folder** object. The destination folder.|

## Return value

An **Object** value that represents the item which has been moved to the designated folder.


## Example

This Visual Basic for Applications (VBA) example uses  **[GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md)** to return the **Folder** object that represents the default folder. It then uses the **[Find](Outlook.Items.Find.md)** and **[FindNext](Outlook.Items.FindNext.md)** methods to find all messages sent by Dan Wilson and uses the **Move** method to move all email messages sent by Dan Wilson from the default **Inbox** folder to the Personal Mail folder. To run this example without any errors, replace 'Dan Wilson' with a vaid sender name and make sure there's a folder under Inbox called 'Personal Mail'. Note that `myItem` is declared as type **Object** so that it can represent all types of Outlook items including meeting request and task request items.


```vb
Sub MoveItems() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myInbox As Outlook.Folder 
 Dim myDestFolder As Outlook.Folder 
 Dim myItems As Outlook.Items 
 Dim myItem As Object 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox) 
 Set myItems = myInbox.Items 
 Set myDestFolder = myInbox.Folders("Personal Mail") 
 Set myItem = myItems.Find("[SenderName] = 'Dan Wilson'") 
 While TypeName(myItem) <> "Nothing" 
 myItem.Move myDestFolder 
 Set myItem = myItems.FindNext 
 Wend 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
