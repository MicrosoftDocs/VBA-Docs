---
title: Calendar object (Project)
ms.prod: project-server
api_name:
- Project.Calendar
ms.assetid: 2d3b0f05-4762-0058-15d4-47e1d2b9d9a9
ms.date: 06/08/2017
localization_priority: Normal
---


# Calendar object (Project)



Represents the calendar for a resource or project. The **Calendar** object is a member of the **[Calendars](Project.calendars.md)** collection.
 **Using the Calendar Object**
Use **BaseCalendars** (_index_), where _index_ is the calendar index number or calendar name, to return a single **Calendar** object.

 **Using the Calendars Collection**
Use the **[BaseCalendars](./Project.Project.BaseCalendars.md)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.
Use the **[BaseCalendarCreate](./Project.Application.BaseCalendarCreate.md)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.

## Methods



|Name|
|:-----|
|[Delete](./Project.Calendar.Delete.md)|
|[Period](./Project.Calendar.Period.md)|
|[Reset](./Project.Calendar.Reset.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.Calendar.Application.md)|
|[BaseCalendar](./Project.Calendar.BaseCalendar.md)|
|[Enterprise](./Project.Calendar.Enterprise.md)|
|[Exceptions](./Project.Calendar.Exceptions.md)|
|[Guid](./Project.Calendar.Guid.md)|
|[Index](./Project.Calendar.Index.md)|
|[Name](./Project.Calendar.Name.md)|
|[Parent](./Project.Calendar.Parent.md)|
|[ResourceGuid](./Project.Calendar.ResourceGuid.md)|
|[WeekDays](./Project.Calendar.WeekDays.md)|
|[WorkWeeks](./Project.Calendar.WorkWeeks.md)|
|[Years](./Project.Calendar.Years.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]