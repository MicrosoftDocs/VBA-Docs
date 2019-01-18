---
title: Customizing the Navigation Pane
ms.prod: outlook
ms.assetid: 426c3d1c-13b5-cac5-702d-87dfe71f2478
ms.date: 06/08/2017
localization_priority: Normal
---


# Customizing the Navigation Pane

The Navigation Pane provides access to information that pertains to the active explorer, including different views and different ways to accomplish tasks in that explorer. The  **[NavigationPane](../../../api/Outlook.NavigationPane.md)** object represents the Navigation Pane for an explorer; to obtain one, call the **[NavigationPane](../../../api/Outlook.Explorer.NavigationPane.md)** property of the **[Explorer](../../../api/Outlook.Explorer.md)** object. If the explorer does not contain a Navigation Pane, this property returns **Null** (**Nothing** in Visual Basic).


## Navigation Modules

The Navigation Pane contains the set of navigation modules that are available in Outlook; for example, the  **Mail** module. Each navigation module is represented by a **[NavigationModule](../../../api/Outlook.NavigationModule.md)** object or by an object that is derived from the **NavigationModule** object. The **[Modules](../../../api/Outlook.NavigationPane.Modules.md)** property of the **NavigationPane** object provides access to the navigation modules that are in the Navigation Pane. You can use the following objects to access the corresponding navigation module:



|**Navigation module**|**Object**|
|:-----|:-----|
| **Calendar**| **[CalendarModule](../../../api/Outlook.CalendarModule.md)**|
| **Contacts**| **[ContactsModule](../../../api/Outlook.ContactsModule.md)**|
| **Journal**| **[JournalModule](../../../api/Outlook.JournalModule.md)**|
| **Folder List**| **NavigationModule**|
| **Mail**| **[MailModule](../../../api/Outlook.MailModule.md)**|
| **Notes**| **[NotesModule](../../../api/Outlook.NotesModule.md)**|
| **Shortcuts**| **NavigationModule**|
| **Solutions**| **[SolutionsModule](../../../api/Outlook.SolutionsModule.md)**|
| **Tasks**| **[TasksModule](../../../api/Outlook.TasksModule.md)**|

Note that the  **Solutions** module is not displayed in the Navigation Pane by default, and can only be created programmatically. The default name of the module is **Solutions**, but you can customize that name.


## Navigation Groups and Navigation Folders

Each navigation module contains a set of navigation groups. A navigation group, represented by the  **[NavigationGroup](../../../api/Outlook.NavigationGroup.md)** object, is a container for navigation folders. A navigation folder, represented by the **[NavigationFolder](../../../api/Outlook.Folder.md)** object, provides an access point in the Navigation Pane for a **[Folder](../../../api/Outlook.Folder.md)** object. You can obtain a **NavigationGroup** object reference by using the **[NavigationGroups](../../../api/Outlook.NavigationGroups.md)** property of a **CalendarModule**,  **ContactsModule**,  **JournalModule**,  **MailModule**,  **NotesModule**, or  **TasksModule** object. The **Folder List**,  **Shortcuts**, and  **Solutions** navigation modules do not contain navigation groups.

You can create and delete custom navigation groups by using the  **[NavigationGroups.Create](../../../api/Outlook.NavigationGroups.Create.md)** and **[NavigationGroups.Delete](../../../api/Outlook.NavigationGroups.Delete.md)** methods. You can identify a custom navigation group by using the **[NavigationGroup.GroupType](../../../api/Outlook.NavigationGroup.GroupType.md)** property to retrieve the navigation group type for the object, and you can retrieve the default navigation group for a specified group type by using the **[NavigationGroups.GetDefaultNavigationGroup](../../../api/Outlook.NavigationGroups.GetDefaultNavigationGroup.md)** method.

Once you have a  **NavigationGroup** object, you can obtain a **NavigationFolder** object reference by using the **[NavigationGroup.NavigationFolders](../../../api/Outlook.NavigationGroup.NavigationFolders.md)** property. Each **NavigationFolder** represents a navigation folder associated with a **Folder** object. You can add navigation folders to a navigation group by using the **[NavigationFolders.Add](../../../api/Outlook.NavigationFolders.Add.md)** method. Only one **NavigationFolder** object can be associated with a **Folder** object at any given time, so adding a **NavigationFolder** that is associated with a given **Folder** object to a navigation group automatically removes any existing **NavigationFolder** references that are associated with that **Folder** object. You can also delete navigation folders from a navigation group by using the **[NavigationFolders.Remove](../../../api/Outlook.NavigationFolders.Remove.md)** method, but only if the **[NavigationFolders.IsRemovable](../../../api/Outlook.NavigationFolder.IsRemovable.md)** property is set to **True** for the **NavigationFolder** object to be removed. You cannot remove standard navigation folders, such as the **Inbox** folder, that are defined by Outlook.


 **Note**  Navigation folders can be freely added or removed from the  **Favorite Folders** navigation group, a special navigation group that is contained by the **MailModule** object, regardless of the **IsRemovable** property value of the navigation folder.


## Displaying the Navigation Pane

The Navigation Pane can display navigation modules in either normal or collapsed mode. The  **[Visible](../../../api/Outlook.NavigationModule.Visible.md)** property of a **NavigationModule** object determines whether the navigation module is displayed in the Navigation Pane; the order that the visible navigation modules are displayed is determined by the **[Position](../../../api/Outlook.NavigationModule.Position.md)** property of each **NavigationModule** object.

You can use the  **[IsCollapsed](../../../api/Outlook.NavigationPane.IsCollapsed.md)** property to determine which mode the **NavigationPane** object uses. In normal mode, the visible navigation modules in the Navigation Pane are displayed as a combination of large and small buttons. The number of large buttons that are displayed in normal mode is determined by the **[DisplayedModuleCount](../../../api/Outlook.NavigationPane.DisplayedModuleCount.md)** property. If there are more visible navigation modules than are specified by this property, the remaining visible navigation modules are displayed as small buttons at the bottom of the Navigation Pane. In collapsed mode, the visible navigation modules in the Navigation Pane are displayed as small buttons. The number of small buttons displayed in collapsed mode is determined by the **DisplayedModuleCount** property. If there are more visible navigation modules than are specified by this property, the remaining visible navigation modules are not displayed.

To change the current navigation module, set the  **[CurrentModule](../../../api/Outlook.NavigationPane.CurrentModule.md)** property of the **NavigationPane** object to one of the **NavigationModule** objects in the navigation pane.

In each navigation module, the  **[NavigationGroup.Position](../../../api/Outlook.NavigationGroup.Position.md)** property determines the display order of the navigation groups. Similarly, the **[NavigationFolder.Position](../../../api/Outlook.NavigationFolder.Position.md)** property determines the display order of navigation folders within each navigation group. If a **NavigationFolder** object represents a calendar folder, the **[IsSideBySide](../../../api/Outlook.NavigationFolder.IsSideBySide.md)** property determines if the contents of the calendar folder are displayed in side-by-side or overlay mode.


## Handling Navigation Pane Events

The  **NavigationPane** object provides the **[ModuleSwitch](../../../api/Outlook.NavigationPane.ModuleSwitch.md)** event so that add-ins can identify when the current navigation module changes in the Navigation Pane, either programmatically or by user action.

The  **NavigationGroups** object provides the **[NavigationFolderAdd](../../../api/Outlook.NavigationGroups.NavigationFolderAdd.md)** and **[NavigationFolderRemove](../../../api/Outlook.NavigationGroups.NavigationFolderRemove.md)** events so that add-ins can identify when a navigation folder is added or removed from a **NavigationGroup** object in the collection. The **NavigationGroups** object also provides the **[SelectedChange](../../../api/Outlook.NavigationGroups.SelectedChange.md)** event. Add-ins use that event to identify when the **[IsSelected](../../../api/Outlook.NavigationFolder.IsSelected.md)** property of a navigation folder that is associated with a calendar folder changes in the Navigation Pane, either programmatically or by user action.

To detect a user change a folder in the Folder List, use the  **[BeforeFolderSwitch](../../../api/Outlook.Explorer.BeforeFolderSwitch.md)** and **[FolderSwitch](../../../api/Outlook.Explorer.FolderSwitch.md)** events of the **[Explorer](../../../api/Outlook.Explorer.md)** object. Similarly, to detect when the **Solutions** module is first displayed in the Navigation Pane, or to detect a user click a different folder in the **Solutions** module, use the **BeforeFolderSwitch** and **FolderSwitch** events.


## See also


 [Adding Solution-Specific Folders to the Solutions Module](adding-solution-specific-folders-to-the-solutions-module.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]