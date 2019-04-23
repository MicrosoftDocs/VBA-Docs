---
title: MailItem.Copy method (Outlook)
keywords: vbaol11.chm1321
f1_keywords:
- vbaol11.chm1321
ms.prod: outlook
api_name:
- Outlook.MailItem.Copy
ms.assetid: a9356844-e31e-eb0f-c0f5-a2923ad127db
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Copy method (Outlook)

Creates another instance of an object.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


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


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
