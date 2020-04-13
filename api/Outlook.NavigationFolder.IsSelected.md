---
title: NavigationFolder.IsSelected property (Outlook)
keywords: vbaol11.chm2905
f1_keywords:
- vbaol11.chm2905
ms.prod: outlook
api_name:
- Outlook.NavigationFolder.IsSelected
ms.assetid: a8fb9430-0477-2417-0dba-e30e9f8ebe8d
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationFolder.IsSelected property (Outlook)

Returns or sets a **Boolean** variable that indicates whether the **[NavigationFolder](Outlook.NavigationFolder.md)** object is selected for display. Read/write.


## Syntax

_expression_. `IsSelected`

_expression_ A variable that represents a [NavigationFolder](Outlook.NavigationFolder.md) object.


## Remarks

Navigation folders contained in a **Calendar** navigation module are treated differently than navigation folders in other navigation modules.

If the active explorer uses the  **Day/Week/Month** or **Day/Week/Month View with AutoPreview** view to display navigation folders in the **Calendar** navigation module, this property returns **True** if the navigation folder is checked in the navigation pane (and is therefore displayed either in side-by-side or overlay mode in the active explorer.) Setting this property to **False** removes a calendar from display in the active explorer. An error occurs if this property is set to **True** for more than 30 navigation folders.

If the active explorer uses another view, such as the  **All Appointments** view, to display navigation folders in the **Calendar** navigation module, or in navigation modules other than the **Calendar** navigation module, this property returns **True** if the navigation folder is selected and currently displayed in the active explorer; otherwise, the property returns **False**. 

In either case, an error occurs if the value of this property is set to  **False** for all **NavigationFolder** objects in the parent **[NavigationFolders](Outlook.NavigationFolders.md)** collection, or if the **NavigationFolder** object is contained by a navigation module other than the navigation module currently displayed in the navigation pane.

The **[SelectedChange](Outlook.NavigationGroups.SelectedChange.md)** event for the parent **NavigationFolders** collection is raised if the value of this property is changed for a **NavigationFolder** object in a **Calendar** navigation module, regardless of the current view.


## See also


[NavigationFolder Object](Outlook.NavigationFolder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]