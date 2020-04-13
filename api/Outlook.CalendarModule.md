---
title: CalendarModule object (Outlook)
keywords: vbaol11.chm3194
f1_keywords:
- vbaol11.chm3194
ms.prod: outlook
api_name:
- Outlook.CalendarModule
ms.assetid: 9203024d-9cef-75e0-600f-f3899e24761a
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarModule object (Outlook)

Represents the  **Calendar** navigation module in the navigation pane of an explorer.


## Remarks

The **CalendarModule** object, derived from the **[NavigationModule](Outlook.NavigationModule.md)** object, provides access to the navigation groups contained in the **Calendar** navigation module of the navigation pane for an explorer. Use the **[GetNavigationModule](Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](Outlook.NavigationModules.Item.md)** method of the **[Modules](Outlook.NavigationPane.Modules.md)** collection for the parent **[NavigationPane](Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleCalendar**, you can then cast the **NavigationModule** object reference as a **CalendarModule** object to access the **[NavigationGroups](Outlook.CalendarModule.NavigationGroups.md)** property for that navigation module.

You can use the  **[Visible](Outlook.CalendarModule.Visible.md)** property to determine if the navigation module is visible and the **[Position](Outlook.CalendarModule.Position.md)** property to return or set the display position of the navigation module within the navigation pane. You can use the **[Name](Outlook.CalendarModule.Name.md)** property to return the display name of the **Calendar** navigation module within the navigation pane.


## Properties



|Name|
|:-----|
|[Application](Outlook.CalendarModule.Application.md)|
|[Class](Outlook.CalendarModule.Class.md)|
|[Name](Outlook.CalendarModule.Name.md)|
|[NavigationGroups](Outlook.CalendarModule.NavigationGroups.md)|
|[NavigationModuleType](Outlook.CalendarModule.NavigationModuleType.md)|
|[Parent](Outlook.CalendarModule.Parent.md)|
|[Position](Outlook.CalendarModule.Position.md)|
|[Session](Outlook.CalendarModule.Session.md)|
|[Visible](Outlook.CalendarModule.Visible.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[CalendarModule Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]