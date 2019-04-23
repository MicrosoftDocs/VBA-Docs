---
title: Folders.FolderRemove event (Outlook)
keywords: vbaol11.chm310
f1_keywords:
- vbaol11.chm310
ms.prod: outlook
api_name:
- Outlook.Folders.FolderRemove
ms.assetid: 9113c4b9-9a18-76a8-3726-7b55fa6e6365
ms.date: 06/08/2017
localization_priority: Normal
---


# Folders.FolderRemove event (Outlook)

Occurs when a folder is removed from the specified  **[Folders](Outlook.Folders.md)** collection.


## Syntax

_expression_. `FolderRemove`

_expression_ A variable that represents a [Folders](Outlook.Folders.md) object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays a warning message when the user tries to a delete a folder in the Inbox. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim myNS As Outlook.NameSpace 
 
Dim WithEvents myFolders As Outlook.Folders 
 
 
 
Sub Initialize_handler() 
 
 Set myNS = Application.GetNamespace("MAPI") 
 
 Set myFolders = myNS.GetDefaultFolder(olFolderInbox).Folders 
 
End Sub 
 
 
 
Private Sub myFolders_FolderRemove() 
 
 MsgBox ("All the items in the folder are deleted as well.") 
 
End Sub
```


## See also


[Folders Object](Outlook.Folders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]