---
title: NavigationFolders Object (Outlook)
keywords: vbaol11.chm3200
f1_keywords:
- vbaol11.chm3200
ms.prod: outlook
api_name:
- Outlook.NavigationFolders
ms.assetid: ecff93b8-0c3f-5f31-5b61-c46d2622d2af
ms.date: 06/08/2017
---


# NavigationFolders Object (Outlook)

Contains a set of  **[NavigationFolder](navigationOutlook.Folder.md)** objects that represent the navigation folders associated with a navigation group.


## Remarks

Use the  **[NavigationFolders](Outlook.NavigationGroup.NavigationFolders.md)** property of the **[NavigationGroup](Outlook.NavigationGroup.md)** object to return a **NavigationFolders** object.

Use the  **[Add](Outlook.NavigationFolders.Add.md)** method to create a new **NavigationFolder** object based on an existing **[Folder](Outlook.Folder.md)** object and add it to the collection. Use the **[Item](Outlook.NavigationFolders.Item.md)** method to return an existing **NavigationFolder** object contained by the **NavigationFolders** collection. Use the **[Remove](Outlook.NavigationFolders.Remove.md)** method from the **[NavigationFolders](Outlook.NavigationFolders.md)** collection of the parent **[NavigationGroup](Outlook.NavigationGroup.md)** object.

Use the  **[NavigationFolderAdd](Outlook.NavigationGroups.NavigationFolderAdd.md)** and **[NavigationFolderRemove](Outlook.NavigationGroups.NavigationFolderRemove.md)** events to detect when a navigation folder is added or removed, respectively, from the **NavigationFolders** object. Use the **[SelectedChange](Outlook.NavigationGroups.SelectedChange.md)** event to detect changes in selection state for navigation folders contained in the **NavigationFolders** object that are based on calendar folders.

Note that if you delete a  **Folder** using **[Folder.Delete](Outlook.Folder.Delete.md)**, the deletion will be reflected automatically in the Navigation Pane and in the **NavigationFolders** collection, but because the synchronization between the actual folders and the Navigation Pane happens asynchronously, this will take a few milliseconds to complete.


## Methods



|**Name**|
|:-----|
|[Add](Outlook.NavigationFolders.Add.md)|
|[Item](Outlook.NavigationFolders.Item.md)|
|[Remove](Outlook.NavigationFolders.Remove.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationFolders.Application.md)|
|[Class](navigationfolders-class-property-outlook.md)|
|[Count](navigationfolders-count-property-outlook.md)|
|[Parent](navigationfolders-parent-property-outlook.md)|
|[Session](navigationfolders-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
