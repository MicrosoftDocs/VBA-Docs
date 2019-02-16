---
title: Views.Add method (Outlook)
keywords: vbaol11.chm547
f1_keywords:
- vbaol11.chm547
ms.prod: outlook
api_name:
- Outlook.Views.Add
ms.assetid: 8005ca2e-8b28-1286-74d1-448f2a168c65
ms.date: 02/16/2019
localization_priority: Normal
---


# Views.Add method (Outlook)

Creates a new view in the **Views** collection.


## Syntax

_expression_.**Add** (_Name_, _ViewType_, _SaveOption_)

_expression_ A variable that represents a **[Views](Outlook.Views.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _ViewType_|Required| **[OlViewType](Outlook.OlViewType.md)**|The type of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option that specifies the permissions of the new view:<ul><li><b>olViewSaveOptionAllFoldersOfType</b> The view can be accessed in all folders of this type.</li><li><b>olViewSaveOptionThisFolderEveryOne</b> The view can be accessed by all users in this folder only.</li><li><b>olViewSaveOptionThisFolderOnlyMe</b> The view can be accessed in this folder only by the user.</li></ul>|

## Return value

A **[View](Outlook.View.md)** object that represents the new view.


## Remarks

If you add a **View** to a **Views** collection of a folder that is not the current folder, you must first save a copy of the **Views** collection object and then add the **View** to this collection object, as shown in the following code sample. This is a work-around for an existing problem that will otherwise cause a call to **[View.Apply](Outlook.View.Apply.md)** for the added **View** to fail.


```vb
Sub CalendarView() 
 Dim calView As Outlook.View 
 Dim vws As Views 
 
 Set Application.ActiveExplorer.CurrentFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Current folder is Inbox; add a View to the Calendar folder which is not the current folder 
 ' Keep a copy of the object for the Views collection for the Calendar 
 Set vws = Application.Session.GetDefaultFolder(olFolderCalendar).Views 
 ' Add the View to this Views collection object 
 Set calView = vws.Add("New Calendar", olCalendarView, olViewSaveOptionThisFolderEveryone) 
 calView.Save 
 ' This Apply call will be fine 
 calView.Apply 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example creates a new view called New Table and stores it in a variable called `objNewView`.


```vb
Sub CreateView() 
 'Creates a new view 
 Dim objName As Outlook.NameSpace 
 Dim objViews As Outlook.Views 
 Dim objNewView As Outlook.View 
 
 Set objName = Application.GetNamespace("MAPI") 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 ViewType:=olTableView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]