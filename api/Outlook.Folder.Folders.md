---
title: Folder.Folders property (Outlook)
keywords: vbaol11.chm1989
f1_keywords:
- vbaol11.chm1989
ms.prod: outlook
api_name:
- Outlook.Folder.Folders
ms.assetid: 41464c32-023e-9079-4f24-51586305325c
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.Folders property (Outlook)

Returns the  **[Folders](Outlook.Folders.md)** collection that represents all the folders contained in the specified **[Folder](Outlook.Folder.md)**. Read-only.


## Syntax

_expression_. `Folders`

_expression_ A variable that represents a [Folder](Outlook.Folder.md) object.


## Remarks

The  **[NameSpace](Outlook.NameSpace.md)** object is the root of all the folders for the given name space.


## Example

This Visual Basic for Applications (VBA) example uses the  **[Folders.Add](Outlook.Folders.Add.md)** method to add the new folder named "My Personal Contacts" to the default **Contacts** folder.


```vb
Sub CreatePersonalContacts() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderContacts) 
 
 Set myNewFolder = myFolder.Folders.Add("My Personal Contacts") 
 
End Sub
```

This VBA example uses the  **Folders.Add** method to add two new folders in the **Tasks** folder. The first folder, "My Notes Folder", will contain note items. The second folder, "My Contacts Folder", will contain contact items. If the folders already exist, a message box will inform the user.




```vb
Sub CreateFolders() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNotesFolder As Outlook.Folder 
 
 Dim myContactFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderTasks) 
 
 On Error GoTo ErrorHandler 
 
 Set myNotesFolder = _ 
 
 myFolder.Folders.Add("My Notes Folder", olFolderNotes) 
 
 Set myContactFolder = _ 
 
 myFolder.Folders.Add("My Contacts Folder", olFolderContacts) 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "Error creating the folder. The folder may already exist." 
 
 Resume Next 
 
End Sub
```


## See also


[Folder Object](Outlook.Folder.md)



[Folder Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
