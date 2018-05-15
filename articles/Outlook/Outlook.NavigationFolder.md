---
title: NavigationFolder Object (Outlook)
keywords: vbaol11.chm3201
f1_keywords:
- vbaol11.chm3201
ms.prod: outlook
api_name:
- Outlook.NavigationFolder
ms.assetid: c8d7aabb-58ba-df5e-ccdc-06f73db7726c
ms.date: 06/08/2017
---


# NavigationFolder Object (Outlook)

Represents a navigation folder displayed in a navigation group of a navigation module in the Navigation Pane.


## Remarks

Use the  **[Item](Outlook.NavigationFolders.Item.md)** method to retrieve a **NavigationFolder** object from the **[NavigationFolders](Outlook.NavigationFolders.md)** collection of the parent **[NavigationGroup](Outlook.NavigationGroup.md)** object. Use the **[Add](Outlook.NavigationFolders.Add.md)** method of the **NavigationFolders** collection to create a new **NavigationFolder** object based on an existing **[Folder](Outlook.Folder.md)** object.

Use the  **[Folder](Outlook.NavigationFolder.Folder.md)** method to return or set the **Folder** object on which the **NavigationFolder** object is based.

Use the  **[IsSelected](Outlook.NavigationFolder.IsSelected.md)** property to determine if the navigation folder is selected and the **[Position](Outlook.NavigationFolder.Position.md)** property to return or set the display position of the navigation folder within the Navigation Pane. You can also use the **[DisplayName](Outlook.NavigationFolder.DisplayName.md)** property to return the display name of the navigation folder within the Navigation Pane.

Use the  **[IsRemovable](Outlook.NavigationFolder.IsRemovable.md)** property to determine if a navigation folder can be removed from the **NavigationFolders** collection and the **[IsSideBySide](Outlook.NavigationFolder.IsSideBySide.md)** property to return or set the viewing mode for a navigation folder associated with a **[CalendarModule](Outlook.CalendarModule.md)** object.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationFolder.Application.md)|
|[Class](Outlook.NavigationFolder.Class.md)|
|[DisplayName](Outlook.NavigationFolder.DisplayName.md)|
|[Folder](Outlook.NavigationFolder.Folder.md)|
|[IsRemovable](Outlook.NavigationFolder.IsRemovable.md)|
|[IsSelected](Outlook.NavigationFolder.IsSelected.md)|
|[IsSideBySide](Outlook.NavigationFolder.IsSideBySide.md)|
|[Parent](navigationfolder-parent-property-outlook.md)|
|[Position](Outlook.NavigationFolder.Position.md)|
|[Session](navigationfolder-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
