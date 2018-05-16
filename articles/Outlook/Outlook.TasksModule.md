---
title: TasksModule Object (Outlook)
keywords: vbaol11.chm3196
f1_keywords:
- vbaol11.chm3196
ms.prod: outlook
api_name:
- Outlook.TasksModule
ms.assetid: fc6ae6c9-6b13-b5f2-9506-c3dbbe709df6
ms.date: 06/08/2017
---


# TasksModule Object (Outlook)

Represents the  **Tasks** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **TasksModule** object, derived from the **[NavigationModule](Outlook.NavigationModule.md)** object, provides access to the navigation groups contained in the **Tasks** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[NavigationModules](Outlook.NavigationModules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleTasks**, you can then cast the **NavigationModule** object reference as a **TasksModule** object to access the **[NavigationGroups](Outlook.TasksModule.NavigationGroups.md)** property for that navigation module.

You can use the  **[Visible](Outlook.TasksModule.Visible.md)** property to determine if the navigation module is visible and the **[Position](Outlook.TasksModule.Position.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](Outlook.TasksModule.Name.md)** property to return the display name of the **Tasks** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.TasksModule.Application.md)|
|[Class](Outlook.TasksModule.Class.md)|
|[Name](Outlook.TasksModule.Name.md)|
|[NavigationGroups](Outlook.TasksModule.NavigationGroups.md)|
|[NavigationModuleType](Outlook.TasksModule.NavigationModuleType.md)|
|[Parent](tasksmodule-parent-property-outlook.md)|
|[Position](Outlook.TasksModule.Position.md)|
|[Session](tasksmodule-session-property-outlook.md)|
|[Visible](Outlook.TasksModule.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
