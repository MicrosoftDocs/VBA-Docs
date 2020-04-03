---
title: NameSpace.GetDefaultFolder method (Outlook)
keywords: vbaol11.chm761
f1_keywords:
- vbaol11.chm761
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetDefaultFolder
ms.assetid: 761b8b53-dd4d-43e4-c8f0-69cefdf0c77a
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.GetDefaultFolder method (Outlook)

Returns a **[Folder](Outlook.Folder.md)** object that represents the default folder of the requested type for the current profile; for example, obtains the default **Calendar** folder for the user who is currently logged on.


## Syntax

_expression_. `GetDefaultFolder`( `_FolderType_` )

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FolderType_|Required| **[OlDefaultFolders](Outlook.OlDefaultFolders.md)**|The type of default folder to return.|

## Return value

A **Folder** object that represents the default folder of the requested type for the current profile.


## Remarks

To return a specific non-default folder, use the  **[Folders](Outlook.Folders.md)** collection.

If the default folder of the requested type does not exist, depending on the type, Outlook may create and return the folder, or may raise an error. For example, if  **olFolderManagedEmail** is specified as the _FolderType_ but the Managed Folders group has not been deployed, Microsoft Outlook raises an error.


## Example

This Visual Basic for Applications (VBA) example uses the  **[CurrentFolder](Outlook.Explorer.CurrentFolder.md)** property to change the displayed folder to the user's default **Calendar** folder.


```vb
Sub ChangeCurrentFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set Application.ActiveExplorer.CurrentFolder = _ 
 
 myNamespace.GetDefaultFolder(olFolderCalendar) 
 
End Sub
```

This VBA example returns the first folder in the Tasks Folders collection.






```vb
Sub DisplayATaskFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myTasks As Outlook.Folder 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myTasks = myNamespace.GetDefaultFolder(olFolderTasks) 
 
 Set myFolder = myTasks.Folders(1) 
 
 myFolder.Display 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
