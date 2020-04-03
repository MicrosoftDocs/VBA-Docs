---
title: Folders.GetFirst method (Outlook)
keywords: vbaol11.chm47
f1_keywords:
- vbaol11.chm47
ms.prod: outlook
api_name:
- Outlook.Folders.GetFirst
ms.assetid: 74757061-2f38-374e-1624-f8df211a711b
ms.date: 06/08/2017
localization_priority: Normal
---


# Folders.GetFirst method (Outlook)

Returns the first object in the  **[Folders](Outlook.Folders.md)** collection.


## Syntax

_expression_. `GetFirst`

_expression_ A variable that represents a [Folders](Outlook.Folders.md) object.


## Return value

A  **[Folder](Outlook.Folder.md)** object that represents the first object contained by the collection.


## Remarks

Returns  **Nothing** if no first object exists, for example, if there are no objects in the collection.To ensure correct operation of the **GetFirst**, **[GetLast](Outlook.Folders.GetLast.md)**, **[GetNext](Outlook.Folders.GetNext.md)**, and **[GetPrevious](Outlook.Folders.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## Example

This Visual Basic for Applications (VBA) example uses the  **GetFirst** method to locate the first folder in the **Contacts** folder and then copies the folder to the Test folder. Before running this example, you need to make sure the necessary folders exist in the default Contacts and Inbox folders.


```vb
Sub CopyItems() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myDestFolder As Outlook.Folder 
 
 Dim mySourceFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myDestFolder = myNameSpace.GetDefaultFolder(olFolderInbox).Folders("Test") 
 
 Set mySourceFolder = myNameSpace.GetDefaultFolder(olFolderContacts).Folders.GetFirst 
 
 Set myNewFolder = mySourceFolder.CopyTo(myDestFolder) 
 
End Sub
```


## See also


[Folders Object](Outlook.Folders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]