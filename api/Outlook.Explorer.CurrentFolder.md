---
title: Explorer.CurrentFolder property (Outlook)
keywords: vbaol11.chm2762
f1_keywords:
- vbaol11.chm2762
ms.prod: outlook
api_name:
- Outlook.Explorer.CurrentFolder
ms.assetid: 75e7f120-28df-0c3b-ec05-bd880621141b
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.CurrentFolder property (Outlook)

Returns or sets a  **[Folder](Outlook.Folder.md)** object that represents the current folder displayed in the explorer. Read/write.


## Syntax

_expression_. `CurrentFolder`

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Remarks

Use this property to change the folder the user is viewing.


## Example

This Visual Basic for Applications (VBA) example uses the  **[CurrentFolder](Outlook.Explorer.CurrentFolder.md)** property to change the displayed folder to the user's **Calendar** folder.


```vb
Sub ChangeCurrentFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set Application.ActiveExplorer.CurrentFolder = _ 
 
 myNamespace.GetDefaultFolder(olFolderCalendar) 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]