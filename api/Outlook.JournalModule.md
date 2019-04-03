---
title: JournalModule object (Outlook)
keywords: vbaol11.chm3197
f1_keywords:
- vbaol11.chm3197
ms.prod: outlook
api_name:
- Outlook.JournalModule
ms.assetid: 5a696d10-8a10-c01d-cf65-f8a65718f120
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalModule object (Outlook)

Represents the  **Journal** navigation module in the navigation pane of an explorer.


## Remarks

The  **JournalModule** object, derived from the **[NavigationModule](Outlook.NavigationModule.md)** object, provides access to the navigation groups contained in the **Journal** navigation module of the navigation pane for an explorer. Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[Modules](Outlook.NavigationPane.Modules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the module type. If the **NavigationModuleType** property is set to **olModuleJournal**, you can then cast the **NavigationModule** object reference as a **JournalModule** object to access the **[NavigationGroups](Outlook.JournalModule.NavigationGroups.md)** property for that navigation module.

You can use the  **[Visible](Outlook.JournalModule.Visible.md)** property to determine if the navigation module is visible and the **[Position](Outlook.JournalModule.Position.md)** property to return or set the display position of the navigation module within the navigation pane. You can use the **[Name](Outlook.JournalModule.Name.md)** property to return the display name of the **Journal** navigation module within the navigation pane.


## Properties



|Name|
|:-----|
|[Application](Outlook.JournalModule.Application.md)|
|[Class](Outlook.JournalModule.Class.md)|
|[Name](Outlook.JournalModule.Name.md)|
|[NavigationGroups](Outlook.JournalModule.NavigationGroups.md)|
|[NavigationModuleType](Outlook.JournalModule.NavigationModuleType.md)|
|[Parent](Outlook.JournalModule.Parent.md)|
|[Position](Outlook.JournalModule.Position.md)|
|[Session](Outlook.JournalModule.Session.md)|
|[Visible](Outlook.JournalModule.Visible.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]