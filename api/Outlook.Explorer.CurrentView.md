---
title: Explorer.CurrentView property (Outlook)
keywords: vbaol11.chm2766
f1_keywords:
- vbaol11.chm2766
ms.prod: outlook
api_name:
- Outlook.Explorer.CurrentView
ms.assetid: 177e6387-9ccb-cb71-bbe5-332c25485848
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.CurrentView property (Outlook)

Returns or sets a  **Variant** representing the current view. Read/write.


## Syntax

_expression_. `CurrentView`

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Remarks

To obtain a  **[View](Outlook.View.md)** object for the view of the current **[Explorer](Outlook.Explorer.md)**, use **Explorer.CurrentView** instead of the **[CurrentView](Outlook.Folder.CurrentView.md)** property of the current **[Folder](Outlook.Folder.md)** object returned by **[Explorer.CurrentFolder](Outlook.Explorer.CurrentFolder.md)**.

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

When this property is set, two events occur:  **[BeforeViewSwitch](Outlook.Explorer.BeforeViewSwitch.md)** occurs before the actual view change takes place and can be used to cancel the change and **[ViewSwitch](Outlook.Explorer.ViewSwitch.md)** takes place after the change is effective.


## Example

The following Visual Basic for Applications (VBA) example sets the current view in the active explorer to messages if the  **Inbox** is displayed.


```vb
Sub ChangeCurrentView() 
 
 Dim myOlExp As Outlook.Explorer 
 
 
 
 Set myOlExp = Application.ActiveExplorer 
 
 If myOlExp.CurrentFolder = "Inbox" Then 
 
 myOlExp.CurrentView = "Messages" 
 
 End If 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]