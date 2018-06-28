---
title: CalendarModule Object (Outlook)
keywords: vbaol11.chm3194
f1_keywords:
- vbaol11.chm3194
ms.prod: outlook
api_name:
- Outlook.CalendarModule
ms.assetid: 9203024d-9cef-75e0-600f-f3899e24761a
ms.date: 06/08/2017
---


# CalendarModule Object (Outlook)

Represents the  **Calendar** navigation module in the Navigation Pane of an explorer.


## Remarks

The  **CalendarModule** object, derived from the **[NavigationModule](../../../api/Outlook.NavigationModule.md)** object, provides access to the navigation groups contained in the **Calendar** navigation module of the Navigation Pane for an explorer. Use the **[GetNavigationModule](../../../api/Outlook.NavigationModules.GetNavigationModule.md)** method or the **[Item](../../../api/Outlook.NavigationModules.Item.md)** method of the **[Modules](../../../api/Outlook.NavigationPane.Modules.md)** collection for the parent **[NavigationPane](../../../api/Outlook.NavigationPane.md)** object to retrieve a **NavigationModule** object, then use the **[NavigationModuleType](../../../api/Outlook.NavigationModule.NavigationModuleType.md)** property of the **NavigationModule** object to retrieve the navigation module type. If the **NavigationModuleType** property is set to **olModuleCalendar**, you can then cast the **NavigationModule** object reference as a **CalendarModule** object to access the **[NavigationGroups](../../../api/Outlook.CalendarModule.NavigationGroups.md)** property for that navigation module.

You can use the  **[Visible](../../../api/Outlook.CalendarModule.Visible.md)** property to determine if the navigation module is visible and the **[Position](../../../api/Outlook.CalendarModule.Position.md)** property to return or set the display position of the navigation module within the Navigation Pane. You can use the **[Name](../../../api/Outlook.CalendarModule.Name.md)** property to return the display name of the **Calendar** navigation module within the Navigation Pane.


## Properties



|**Name**|
|:-----|
|[Application](../../../api/Outlook.CalendarModule.Application.md)|
|[Class](../../../api/Outlook.CalendarModule.Class.md)|
|[Name](../../../api/Outlook.CalendarModule.Name.md)|
|[NavigationGroups](../../../api/Outlook.CalendarModule.NavigationGroups.md)|
|[NavigationModuleType](../../../api/Outlook.CalendarModule.NavigationModuleType.md)|
|[Parent](../../../api/Outlook.CalendarModule.Parent.md)|
|[Position](../../../api/Outlook.CalendarModule.Position.md)|
|[Session](../../../api/Outlook.CalendarModule.Session.md)|
|[Visible](../../../api/Outlook.CalendarModule.Visible.md)|

## See also


#### Other resources


[Outlook Object Model Reference](../../../api/overview/object-model-outlook-vba-reference.md)
[CalendarModule Object Members](../../../api/overview/Outlook.md)
