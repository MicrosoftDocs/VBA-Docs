---
title: Folder.CurrentView property (Outlook)
keywords: vbaol11.chm2009
f1_keywords:
- vbaol11.chm2009
ms.prod: outlook
api_name:
- Outlook.Folder.CurrentView
ms.assetid: 42af4345-60f1-10cd-66e5-517ca002284b
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.CurrentView property (Outlook)

Returns a  **[View](Outlook.View.md)** object representing the current view. Read-only.


## Syntax

_expression_. `CurrentView`

_expression_ A variable that represents a [Folder](Outlook.Folder.md) object.


## Remarks

To obtain a  **View** object for the view of the current **[Explorer](Outlook.Explorer.md)**, use **[Explorer.CurrentView](Outlook.Explorer.CurrentView.md)** instead of the **CurrentView** property of the current **[Folder](Outlook.Folder.md)** object returned by **[Explorer.CurrentFolder](Outlook.Explorer.CurrentFolder.md)**.

You must save a reference to the  **View** object returned by **CurrentView** before you proceed to use it for any purpose.

To properly reset the current view, you must do a  **[View.Reset](Outlook.View.Reset.md)** and then a **[View.Apply](Outlook.View.Apply.md)**. The code sample below illustrates the order of the calls:




```vb
Sub ResetView() 
 
 Dim v as Outlook.View 
 
 ' Save a reference to the current view object 
 
 Set v = Application.ActiveExplorer.CurrentView 
 
 ' Reset and then apply the current view 
 
 v.Reset 
 
 v.Apply 
 
End Sub
```


## Example

The following VBA example displays the current view of the Inbox folder.


```vb
Sub TestFolderCurrentView() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim mpFolder As Outlook.Folder 
 
 Dim vw As Outlook.View 
 
 Dim strView As String 
 
 
 
 Set nsp = Application.Session 
 
 Set mpFolder = nsp.GetDefaultFolder(olFolderInbox) 
 
 ' Save a reference to the current view 
 
 Set vw = mpFolder.CurrentView 
 
 MsgBox "The Current View is: " & vw.Name 
 
End Sub
```


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]