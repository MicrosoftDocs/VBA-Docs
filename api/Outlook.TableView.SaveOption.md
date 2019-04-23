---
title: TableView.SaveOption property (Outlook)
keywords: vbaol11.chm2511
f1_keywords:
- vbaol11.chm2511
ms.prod: outlook
api_name:
- Outlook.TableView.SaveOption
ms.assetid: ddd50cb7-60e4-e820-3f3a-e84320fc76be
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView.SaveOption property (Outlook)

Returns an  **[OlViewSaveOption](Outlook.OlViewSaveOption.md)** constant that specifies the folders in which the specified view is available and the read permissions attached to the view. Read-only.


## Syntax

_expression_. `SaveOption`

_expression_ A variable that represents a [TableView](Outlook.TableView.md) object.


## Remarks

The value of the  **SaveOption** property is set when the **[TableView](Outlook.tableView.md)** object is created using the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection.


## Example

The following Visual Basic for Applications (VBA) example locks the user interface for all views that are available to all users. The subroutine  `LockView` accepts the **[View](Outlook.View.md)** object and a **Boolean** value that indicates if the **View** user interface will be locked. In this example, the procedure is always called with the **Boolean** value set to **True**.


```vb
Sub LockPublicViews() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 ' Get the Views collection for the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Enumerate the Views collection and lock the user 
 
 ' interface for any view that can be accessed by 
 
 ' all users who have access to the Notes default folder. 
 
 For Each objView In objViews 
 
 If objView.SaveOption = _ 
 
 olViewSaveOptionThisFolderEveryone Then 
 
 
 
 Call LockView(objView, True) 
 
 End If 
 
 Next objView 
 
 
 
End Sub 
 
 
 
Sub LockView(ByRef objView As View, ByVal blnAns As Boolean) 
 
 
 
 ' Examine the view object. 
 
 With objView 
 
 If blnAns = True Then 
 
 ' Lock the user interface and 
 
 ' save the view 
 
 .LockUserChanges = True 
 
 .Save 
 
 Else 
 
 ' Unlock the user interface of the view. 
 
 .LockUserChanges = False 
 
 End If 
 
 End With 
 
 
 
End Sub
```


## See also


[TableView Object](Outlook.tableView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]