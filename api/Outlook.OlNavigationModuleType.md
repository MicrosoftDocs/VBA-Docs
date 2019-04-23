---
title: OlNavigationModuleType enumeration (Outlook)
keywords: vbaol11.chm3145
f1_keywords:
- vbaol11.chm3145
ms.prod: outlook
api_name:
- Outlook.OlNavigationModuleType
ms.assetid: 2140a094-6bee-aba1-03cd-71fa2c55842e
ms.date: 06/08/2017
localization_priority: Normal
---


# OlNavigationModuleType enumeration (Outlook)

Identifies the navigation module type of a  **[NavigationModule](Outlook.NavigationModule.md)** object.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olModuleCalendar**|1|A  **[CalendarModule](Outlook.CalendarModule.md)** object that represents the **Calendar** navigation module.|
| **olModuleContacts**|2|A  **[ContactsModule](Outlook.ContactsModule.md)** object that represents the **Contacts** navigation module.|
| **olModuleFolderList**|6|A  **NavigationModule** object that represents the **Folders List** navigation module.|
| **olModuleJournal**|4|A  **[JournalModule](Outlook.JournalModule.md)** object that represents the **Journal** navigation module.|
| **olModuleMail**|0|A  **[MailModule](Outlook.MailModule.md)** object that represents the **Mail** navigation module.|
| **olModuleNotes**|5|A  **[NotesModule](Outlook.NotesModule.md)** object that represents the **Notes** navigation module.|
| **olModuleShortcuts**|7|A  **NavigationModule** object that represents the **Shortcuts** navigation module.|
| **olModuleSolutions**|8|A  **[SolutionsModule](Outlook.SolutionsModule.md)** object that represents the **Solutions** navigation module.|
| **olModuleTasks**|3|A  **[TasksModule](Outlook.TasksModule.md)** object that represents the **Tasks** navigation module.|

## Remarks

This enumeration is used by the [NavigationModule.NavigationModuleType property (Outlook)](Outlook.NavigationModule.NavigationModuleType.md) for the following objects to identify the type of navigation module:


1.  **CalendarModule**
    
2.  **ContactsModule**
    
3.  **JournalModule**
    
4.  **MailModule**
    
5.  **Module**
    
6.  **NotesModule**
    
7.  **SolutionsModule**
    
8.  **TasksModule**
    
The enumeration is also used by the [NavigationModules.GetNavigationModule method (Outlook)](Outlook.NavigationModules.GetNavigationModule.md) to identify the navigation module type of the **NavigationModule** to retrieve.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]