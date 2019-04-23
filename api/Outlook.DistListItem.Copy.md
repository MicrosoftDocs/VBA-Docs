---
title: DistListItem.Copy method (Outlook)
keywords: vbaol11.chm1141
f1_keywords:
- vbaol11.chm1141
ms.prod: outlook
api_name:
- Outlook.DistListItem.Copy
ms.assetid: c6d41be4-9c8f-a25e-9379-ef24876b6e9d
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.Copy method (Outlook)

Creates another instance of an object.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Example

This Visual Basic for Applications example creates an email message, sets the  **Subject** to "Speeches", uses the **Copy** method to copy it, then moves the copy into a newly created email folder named "Saved Mail" within the Inbox folder.


```vb
Sub CopyItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myCopiedItem As Outlook.MailItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myNewFolder = myFolder.Folders.Add("Saved Mail", olFolderDrafts) 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Speeches" 
 
 Set myCopiedItem = myItem.Copy 
 
 myCopiedItem.Move myNewFolder 
 
End Sub
```


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]