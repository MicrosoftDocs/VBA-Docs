---
title: NavigationModule Object (Outlook)
keywords: vbaol11.chm3211
f1_keywords:
- vbaol11.chm3211
ms.prod: outlook
api_name:
- Outlook.NavigationModule
ms.assetid: 76565eaf-1e64-f5d4-b90f-ba156863802c
ms.date: 06/08/2017
---


# NavigationModule Object (Outlook)

Represents a navigation module in the Navigation Pane.


## Remarks

The  **NavigationModule** object provides access to the various navigation modules that are displayed in the Microsoft Outlook Navigation Pane. The following objects are derived from the **NavigationModule** object:


-  **[CalendarModule](Outlook.CalendarModule.md)**
    
-  **[ContactsModule](Outlook.ContactsModule.md)**
    
-  **[JournalModule](Outlook.JournalModule.md)**
    
-  **[MailModule](Outlook.MailModule.md)**
    
-  **[NotesModule](Outlook.NotesModule.md)**
    
-  **[TasksModule](Outlook.TasksModule.md)**
    
-  **[SolutionsModule](Outlook.SolutionsModule.md)**
    
 Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[NavigationModules](Outlook.NavigationModules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the module type. Depending on the value of the **NavigationModuleType** property, you can then cast the **NavigationModule** object reference as one of the objects listed in the previous paragraph to access the **[NavigationGroups](Outlook.CalendarModule.NavigationGroups.md)** property for that object, such as a **CalendarModule** object.

The  **Shortcuts** and **Folder List** navigation modules do not have a corresponding object, such as **MailModule**, because they do not support programmatic access to navigation groups or navigation folders. You can use the **NavigationModule** object to access the properties of the **Shortcuts** and **Folder List** modules.

You can use the  **[Visible](Outlook.NavigationModule.Visible.md)** property to determine whether the navigation module is visible, and use the **[Position](Outlook.NavigationModule.Position.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can also use the **[Name](Outlook.NavigationModule.Name.md)** property to return the display name of the navigation module in the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](Outlook.NavigationModule.Application.md)|
|[Class](Outlook.NavigationModule.Class.md)|
|[Name](Outlook.NavigationModule.Name.md)|
|[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)|
|[Parent](navigationmodule-parent-property-outlook.md)|
|[Position](Outlook.NavigationModule.Position.md)|
|[Session](navigationmodule-session-property-outlook.md)|
|[Visible](Outlook.NavigationModule.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
